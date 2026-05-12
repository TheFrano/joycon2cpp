#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/Windows.Storage.Streams.h>
#include <winrt/Windows.Devices.Bluetooth.h>
#include <winrt/Windows.Devices.Bluetooth.Advertisement.h>
#include <winrt/Windows.Devices.Bluetooth.GenericAttributeProfile.h>
#pragma comment(lib, "setupapi.lib")
#include <iostream>
#include <vector>
#include <algorithm>
#include <thread>
#include <mutex>
#include <atomic>
#include <condition_variable>
#include <memory>
#include <fstream>
#include <sstream>
#include <map>
#include <chrono>
#include <iomanip>
#include <string>
#include <conio.h>  // For _kbhit()
#include "JoyConDecoder.h"
#include "DsuServer.h"
#include <Windows.h>

#include <ViGEm/Client.h>
#include <ViGEm/Common.h>

using namespace winrt;
using namespace Windows::Devices::Bluetooth;
using namespace Windows::Devices::Bluetooth::Advertisement;
using namespace Windows::Devices::Bluetooth::GenericAttributeProfile;
using namespace Windows::Storage::Streams;
using namespace Windows::Foundation;

constexpr uint16_t JOYCON_MANUFACTURER_ID = 1363; // Nintendo
const std::vector<uint8_t> JOYCON_MANUFACTURER_PREFIX = { 0x01, 0x00, 0x03, 0x7E };
const wchar_t* INPUT_REPORT_UUID = L"ab7de9be-89fe-49ad-828f-118f09df7fd2";
const wchar_t* WRITE_COMMAND_UUID = L"649d4ac9-8eb7-4e6c-af44-1ea54fe5f005";

// GL/GR Button Mapping Configuration
enum class ButtonMapping {
    NONE,
    L3,        // Left stick click
    R3,        // Right stick click
    L1,        // Left shoulder
    R1,        // Right shoulder
    L2,        // Left trigger
    R2,        // Right trigger
    CROSS,     // X / A
    CIRCLE,    // O / B
    SQUARE,    // □ / X
    TRIANGLE,  // △ / Y
    SHARE,     // Share / Back
    OPTIONS,   // Options / Start
    DPAD_UP,
    DPAD_DOWN,
    DPAD_LEFT,
    DPAD_RIGHT
};

// Single GL/GR mapping layout
struct GLGRLayout {
    std::string name;
    ButtonMapping glMapping;
    ButtonMapping grMapping;
};

// Pro Controller configuration with multiple layouts
struct ProControllerConfig {
    std::vector<GLGRLayout> layouts;
    int activeLayoutIndex = 0;
};

const std::string CONFIG_FILE = "joycon2cpp_config.json";

// Global config instance
ProControllerConfig g_proControllerConfig;

PVIGEM_CLIENT vigem_client = nullptr;

using SteadyClock = std::chrono::steady_clock;
using TimePoint = SteadyClock::time_point;

enum class UpdatePolicy {
    LowLatency,
    Balanced120Hz,
    Legacy60Hz
};

const char* UpdatePolicyName(UpdatePolicy policy)
{
    switch (policy) {
        case UpdatePolicy::LowLatency: return "LowLatency";
        case UpdatePolicy::Balanced120Hz: return "Balanced120Hz";
        case UpdatePolicy::Legacy60Hz: return "Legacy60Hz";
        default: return "Unknown";
    }
}

const char* ControllerTypeName(int controllerType)
{
    switch (controllerType) {
        case 1: return "SingleJoyCon";
        case 2: return "DualJoyCon";
        case 3: return "ProController";
        case 4: return "NSOGCController";
        default: return "Unknown";
    }
}

struct RuntimeOptions {
    bool latencyMetrics = false;
    UpdatePolicy updatePolicy = UpdatePolicy::LowLatency;
    std::string latencyCsvPath = "latency_benchmark.csv";
};

RuntimeOptions g_runtimeOptions;

class LatencyCsvLogger {
public:
    bool Start(const std::string& path)
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        m_file.open(path, std::ios::out | std::ios::trunc);
        if (!m_file.is_open()) {
            return false;
        }

        m_file << "mode,controller_type,event_index,ble_delta_ms,buffer_age_left_ms,buffer_age_right_ms,decode_to_vigem_us,total_pipeline_us\n";
        m_enabled = true;
        return true;
    }

    bool Enabled() const
    {
        return m_enabled.load(std::memory_order_acquire);
    }

    void Record(
        UpdatePolicy policy,
        const char* controllerType,
        uint64_t eventIndex,
        double bleDeltaMs,
        double bufferAgeLeftMs,
        double bufferAgeRightMs,
        double decodeToVigemUs,
        double totalPipelineUs)
    {
        if (!Enabled()) {
            return;
        }

        std::lock_guard<std::mutex> lock(m_mutex);
        if (!m_file.is_open()) {
            return;
        }

        m_file << UpdatePolicyName(policy) << ','
               << controllerType << ','
               << eventIndex << ','
               << std::fixed << std::setprecision(3)
               << bleDeltaMs << ','
               << bufferAgeLeftMs << ','
               << bufferAgeRightMs << ','
               << std::setprecision(1)
               << decodeToVigemUs << ','
               << totalPipelineUs << '\n';
    }

private:
    std::atomic<bool> m_enabled{ false };
    std::mutex m_mutex;
    std::ofstream m_file;
};

LatencyCsvLogger g_latencyLogger;

double MillisecondsBetween(TimePoint start, TimePoint end)
{
    if (start == TimePoint{}) {
        return -1.0;
    }
    return std::chrono::duration<double, std::milli>(end - start).count();
}

double MicrosecondsBetween(TimePoint start, TimePoint end)
{
    return std::chrono::duration<double, std::micro>(end - start).count();
}

std::chrono::microseconds PolicyInterval(UpdatePolicy policy)
{
    switch (policy) {
        case UpdatePolicy::Balanced120Hz:
            return std::chrono::microseconds(8333);
        case UpdatePolicy::Legacy60Hz:
            return std::chrono::microseconds(16667);
        case UpdatePolicy::LowLatency:
        default:
            return std::chrono::microseconds(0);
    }
}

bool ShouldEmitForPolicy(UpdatePolicy policy, TimePoint& lastEmit, TimePoint now)
{
    const auto interval = PolicyInterval(policy);
    if (interval.count() == 0 || lastEmit == TimePoint{} || now - lastEmit >= interval) {
        lastEmit = now;
        return true;
    }
    return false;
}

struct LatencyTracker {
    TimePoint lastBleTime{};
    TimePoint lastEmitTime{};
    uint64_t eventIndex = 0;
};

void InitializeViGEm()
{
    if (vigem_client != nullptr)
        return;

    vigem_client = vigem_alloc();
    if (vigem_client == nullptr)
    {
        std::wcerr << L"Failed to allocate ViGEm client.\n";
        exit(1);
    }

    auto ret = vigem_connect(vigem_client);
    if (!VIGEM_SUCCESS(ret))
    {
        std::wcerr << L"Failed to connect to ViGEm bus: 0x" << std::hex << ret << L"\n";
        exit(1);
    }

    std::wcout << L"ViGEm client initialized and connected.\n";
}

void PrintRawNotification(const std::vector<uint8_t>& buffer)
{
    std::cout << "[Raw Notification] ";
    for (auto b : buffer) {
        printf("%02X ", b);
    }
    std::cout << std::endl;
}

// Helper: Convert ButtonMapping to string
std::string ButtonMappingToString(ButtonMapping mapping) {
    switch (mapping) {
        case ButtonMapping::NONE: return "NONE";
        case ButtonMapping::L3: return "L3";
        case ButtonMapping::R3: return "R3";
        case ButtonMapping::L1: return "L1";
        case ButtonMapping::R1: return "R1";
        case ButtonMapping::L2: return "L2";
        case ButtonMapping::R2: return "R2";
        case ButtonMapping::CROSS: return "CROSS";
        case ButtonMapping::CIRCLE: return "CIRCLE";
        case ButtonMapping::SQUARE: return "SQUARE";
        case ButtonMapping::TRIANGLE: return "TRIANGLE";
        case ButtonMapping::SHARE: return "SHARE";
        case ButtonMapping::OPTIONS: return "OPTIONS";
        case ButtonMapping::DPAD_UP: return "DPAD_UP";
        case ButtonMapping::DPAD_DOWN: return "DPAD_DOWN";
        case ButtonMapping::DPAD_LEFT: return "DPAD_LEFT";
        case ButtonMapping::DPAD_RIGHT: return "DPAD_RIGHT";
        default: return "NONE";
    }
}

// Helper: Convert string to ButtonMapping
ButtonMapping StringToButtonMapping(const std::string& str) {
    if (str == "L3") return ButtonMapping::L3;
    if (str == "R3") return ButtonMapping::R3;
    if (str == "L1") return ButtonMapping::L1;
    if (str == "R1") return ButtonMapping::R1;
    if (str == "L2") return ButtonMapping::L2;
    if (str == "R2") return ButtonMapping::R2;
    if (str == "CROSS") return ButtonMapping::CROSS;
    if (str == "CIRCLE") return ButtonMapping::CIRCLE;
    if (str == "SQUARE") return ButtonMapping::SQUARE;
    if (str == "TRIANGLE") return ButtonMapping::TRIANGLE;
    if (str == "SHARE") return ButtonMapping::SHARE;
    if (str == "OPTIONS") return ButtonMapping::OPTIONS;
    if (str == "DPAD_UP") return ButtonMapping::DPAD_UP;
    if (str == "DPAD_DOWN") return ButtonMapping::DPAD_DOWN;
    if (str == "DPAD_LEFT") return ButtonMapping::DPAD_LEFT;
    if (str == "DPAD_RIGHT") return ButtonMapping::DPAD_RIGHT;
    return ButtonMapping::NONE;
}

// Simple JSON serialization for our config
std::string ConfigToJSON(const ProControllerConfig& config) {
    std::stringstream ss;
    ss << "{\n";
    ss << "  \"activeLayoutIndex\": " << config.activeLayoutIndex << ",\n";
    ss << "  \"layouts\": [\n";

    for (size_t i = 0; i < config.layouts.size(); ++i) {
        const auto& layout = config.layouts[i];
        ss << "    {\n";
        ss << "      \"name\": \"" << layout.name << "\",\n";
        ss << "      \"glMapping\": \"" << ButtonMappingToString(layout.glMapping) << "\",\n";
        ss << "      \"grMapping\": \"" << ButtonMappingToString(layout.grMapping) << "\"\n";
        ss << "    }";
        if (i < config.layouts.size() - 1) ss << ",";
        ss << "\n";
    }

    ss << "  ]\n";
    ss << "}\n";
    return ss.str();
}

// Simple JSON parsing for our config
bool JSONToConfig(const std::string& json, ProControllerConfig& config) {
    config.layouts.clear();
    config.activeLayoutIndex = 0;

    // Simple parser for our specific JSON structure
    size_t pos = 0;

    // Find activeLayoutIndex
    size_t activePos = json.find("\"activeLayoutIndex\"");
    if (activePos != std::string::npos) {
        size_t colonPos = json.find(':', activePos);
        size_t commaPos = json.find_first_of(",\n", colonPos);
        std::string value = json.substr(colonPos + 1, commaPos - colonPos - 1);
        // Trim whitespace
        value.erase(0, value.find_first_not_of(" \t\n\r"));
        value.erase(value.find_last_not_of(" \t\n\r") + 1);
        config.activeLayoutIndex = std::stoi(value);
    }

    // Find layouts array
    size_t layoutsStart = json.find("\"layouts\"");
    if (layoutsStart == std::string::npos) return false;

    size_t arrayStart = json.find('[', layoutsStart);
    if (arrayStart == std::string::npos) return false;

    // Parse each layout object
    pos = arrayStart + 1;
    while (pos < json.length()) {
        size_t objectStart = json.find('{', pos);
        if (objectStart == std::string::npos) break;

        size_t objectEnd = json.find('}', objectStart);
        if (objectEnd == std::string::npos) break;

        std::string objectStr = json.substr(objectStart, objectEnd - objectStart + 1);

        GLGRLayout layout;

        // Parse name
        size_t namePos = objectStr.find("\"name\"");
        if (namePos != std::string::npos) {
            size_t nameStart = objectStr.find('\"', namePos + 6);
            size_t nameEnd = objectStr.find('\"', nameStart + 1);
            layout.name = objectStr.substr(nameStart + 1, nameEnd - nameStart - 1);
        }

        // Parse glMapping
        size_t glPos = objectStr.find("\"glMapping\"");
        if (glPos != std::string::npos) {
            size_t glStart = objectStr.find('\"', glPos + 11);
            size_t glEnd = objectStr.find('\"', glStart + 1);
            std::string glStr = objectStr.substr(glStart + 1, glEnd - glStart - 1);
            layout.glMapping = StringToButtonMapping(glStr);
        }

        // Parse grMapping
        size_t grPos = objectStr.find("\"grMapping\"");
        if (grPos != std::string::npos) {
            size_t grStart = objectStr.find('\"', grPos + 11);
            size_t grEnd = objectStr.find('\"', grStart + 1);
            std::string grStr = objectStr.substr(grStart + 1, grEnd - grStart - 1);
            layout.grMapping = StringToButtonMapping(grStr);
        }

        config.layouts.push_back(layout);

        pos = objectEnd + 1;
        // Check if there's another object
        size_t nextComma = json.find(',', pos);
        size_t arrayEnd = json.find(']', pos);
        if (arrayEnd != std::string::npos && (nextComma == std::string::npos || arrayEnd < nextComma)) {
            break;
        }
    }

    return !config.layouts.empty();
}

// Load config from JSON file
bool LoadProControllerConfig(ProControllerConfig& config) {
    std::ifstream file(CONFIG_FILE);
    if (!file.is_open()) {
        return false;
    }

    std::stringstream buffer;
    buffer << file.rdbuf();
    file.close();

    return JSONToConfig(buffer.str(), config);
}

// Save config to JSON file
void SaveProControllerConfig(const ProControllerConfig& config) {
    std::ofstream file(CONFIG_FILE);
    if (!file.is_open()) {
        std::cerr << "Failed to save config file.\n";
        return;
    }

    file << ConfigToJSON(config);
    file.close();
    std::cout << "Configuration saved to " << CONFIG_FILE << "\n";
}

// Prompt user for button mapping
ButtonMapping PromptForButtonMapping(const std::string& buttonName) {
    std::cout << "\nSelect mapping for " << buttonName << " button:\n";
    std::cout << "  1. L3 (Left Stick Click)\n";
    std::cout << "  2. R3 (Right Stick Click)\n";
    std::cout << "  3. L1 (Left Shoulder)\n";
    std::cout << "  4. R1 (Right Shoulder)\n";
    std::cout << "  5. L2 (Left Trigger)\n";
    std::cout << "  6. R2 (Right Trigger)\n";
    std::cout << "  7. Cross (X/A)\n";
    std::cout << "  8. Circle (O/B)\n";
    std::cout << "  9. Square (□/X)\n";
    std::cout << " 10. Triangle (△/Y)\n";
    std::cout << " 11. Share (Back)\n";
    std::cout << " 12. Options (Start)\n";
    std::cout << " 13. D-Pad Up\n";
    std::cout << " 14. D-Pad Down\n";
    std::cout << " 15. D-Pad Left\n";
    std::cout << " 16. D-Pad Right\n";
    std::cout << " 17. None (Disable)\n";
    std::cout << "Enter choice (1-17): ";

    int choice;
    std::cin >> choice;
    std::cin.ignore((std::numeric_limits<std::streamsize>::max)(), '\n');

    switch (choice) {
        case 1: return ButtonMapping::L3;
        case 2: return ButtonMapping::R3;
        case 3: return ButtonMapping::L1;
        case 4: return ButtonMapping::R1;
        case 5: return ButtonMapping::L2;
        case 6: return ButtonMapping::R2;
        case 7: return ButtonMapping::CROSS;
        case 8: return ButtonMapping::CIRCLE;
        case 9: return ButtonMapping::SQUARE;
        case 10: return ButtonMapping::TRIANGLE;
        case 11: return ButtonMapping::SHARE;
        case 12: return ButtonMapping::OPTIONS;
        case 13: return ButtonMapping::DPAD_UP;
        case 14: return ButtonMapping::DPAD_DOWN;
        case 15: return ButtonMapping::DPAD_LEFT;
        case 16: return ButtonMapping::DPAD_RIGHT;
        case 17: return ButtonMapping::NONE;
        default:
            std::cout << "Invalid choice, defaulting to NONE.\n";
            return ButtonMapping::NONE;
    }
}

// Configure GL/GR mappings interactively
// Create default config with one empty layout
void CreateDefaultConfig() {
    g_proControllerConfig.layouts.clear();
    GLGRLayout defaultLayout;
    defaultLayout.name = "Layout 1";
    defaultLayout.glMapping = ButtonMapping::NONE;
    defaultLayout.grMapping = ButtonMapping::NONE;
    g_proControllerConfig.layouts.push_back(defaultLayout);
    g_proControllerConfig.activeLayoutIndex = 0;
    SaveProControllerConfig(g_proControllerConfig);
}

// Configure a single layout
void ConfigureLayout(GLGRLayout& layout) {
    std::cout << "\n=== Configure GL/GR Mapping ===\n";
    std::cout << "Layout Name: " << layout.name << "\n\n";

    layout.glMapping = PromptForButtonMapping("GL");
    layout.grMapping = PromptForButtonMapping("GR");

    std::cout << "\nLayout configured!\n";
    std::cout << "  GL -> " << ButtonMappingToString(layout.glMapping) << "\n";
    std::cout << "  GR -> " << ButtonMappingToString(layout.grMapping) << "\n";
}

// Layout management window
void ShowLayoutManagementWindow() {
    while (true) {
        std::cout << "\n==================================================\n";
        std::cout << "          GL/GR LAYOUT MANAGEMENT\n";
        std::cout << "==================================================\n";
        std::cout << "Current active layout: " << g_proControllerConfig.layouts[g_proControllerConfig.activeLayoutIndex].name << "\n\n";
        std::cout << "Available Layouts:\n";

        for (size_t i = 0; i < g_proControllerConfig.layouts.size(); ++i) {
            const auto& layout = g_proControllerConfig.layouts[i];
            std::cout << "  " << (i + 1) << ". " << layout.name;
            if (i == static_cast<size_t>(g_proControllerConfig.activeLayoutIndex)) {
                std::cout << " [ACTIVE]";
            }
            std::cout << "\n";
            std::cout << "     GL: " << ButtonMappingToString(layout.glMapping)
                      << " | GR: " << ButtonMappingToString(layout.grMapping) << "\n";
        }

        std::cout << "  " << (g_proControllerConfig.layouts.size() + 1) << ". [NEW]\n";
        std::cout << "  0. Exit Management Window\n\n";
        std::cout << "Enter number to edit layout: ";

        std::string input;
        std::getline(std::cin, input);

        if (input.empty()) continue;

        int choice = std::stoi(input);

        if (choice == 0) {
            SaveProControllerConfig(g_proControllerConfig);
            std::cout << "Exiting layout management...\n";
            break;
        }
        else if (choice == static_cast<int>(g_proControllerConfig.layouts.size() + 1)) {
            // Create new layout
            GLGRLayout newLayout;
            std::cout << "\nEnter name for new layout: ";
            std::getline(std::cin, newLayout.name);
            if (newLayout.name.empty()) {
                newLayout.name = "Layout " + std::to_string(g_proControllerConfig.layouts.size() + 1);
            }
            newLayout.glMapping = ButtonMapping::NONE;
            newLayout.grMapping = ButtonMapping::NONE;

            ConfigureLayout(newLayout);
            g_proControllerConfig.layouts.push_back(newLayout);
            SaveProControllerConfig(g_proControllerConfig);
            std::cout << "\nNew layout added!\n";
        }
        else if (choice > 0 && choice <= static_cast<int>(g_proControllerConfig.layouts.size())) {
            // Edit existing layout
            size_t index = choice - 1;
            std::cout << "\nEditing: " << g_proControllerConfig.layouts[index].name << "\n";
            std::cout << "1. Rename layout\n";
            std::cout << "2. Configure mappings\n";
            std::cout << "3. Delete layout\n";
            std::cout << "4. Set as active layout\n";
            std::cout << "0. Cancel\n";
            std::cout << "Choice: ";

            std::string editChoice;
            std::getline(std::cin, editChoice);

            if (editChoice == "1") {
                std::cout << "Enter new name: ";
                std::string newName;
                std::getline(std::cin, newName);
                if (!newName.empty()) {
                    g_proControllerConfig.layouts[index].name = newName;
                    SaveProControllerConfig(g_proControllerConfig);
                }
            }
            else if (editChoice == "2") {
                ConfigureLayout(g_proControllerConfig.layouts[index]);
                SaveProControllerConfig(g_proControllerConfig);
            }
            else if (editChoice == "3") {
                if (g_proControllerConfig.layouts.size() > 1) {
                    std::cout << "Are you sure you want to delete this layout? (y/n): ";
                    std::string confirm;
                    std::getline(std::cin, confirm);
                    if (confirm == "y" || confirm == "Y") {
                        g_proControllerConfig.layouts.erase(g_proControllerConfig.layouts.begin() + index);
                        if (g_proControllerConfig.activeLayoutIndex >= static_cast<int>(g_proControllerConfig.layouts.size())) {
                            g_proControllerConfig.activeLayoutIndex = static_cast<int>(g_proControllerConfig.layouts.size()) - 1;
                        }
                        SaveProControllerConfig(g_proControllerConfig);
                        std::cout << "Layout deleted!\n";
                    }
                } else {
                    std::cout << "Cannot delete the last layout!\n";
                }
            }
            else if (editChoice == "4") {
                g_proControllerConfig.activeLayoutIndex = static_cast<int>(index);
                SaveProControllerConfig(g_proControllerConfig);
                std::cout << "Active layout changed to: " << g_proControllerConfig.layouts[index].name << "\n";
            }
        }
    }
}

// Apply ButtonMapping to DS4 report
void ApplyButtonMapping(DS4_REPORT_EX& report, ButtonMapping mapping) {
    switch (mapping) {
        case ButtonMapping::L3:
            report.Report.wButtons |= DS4_BUTTON_THUMB_LEFT;
            break;
        case ButtonMapping::R3:
            report.Report.wButtons |= DS4_BUTTON_THUMB_RIGHT;
            break;
        case ButtonMapping::L1:
            report.Report.wButtons |= DS4_BUTTON_SHOULDER_LEFT;
            break;
        case ButtonMapping::R1:
            report.Report.wButtons |= DS4_BUTTON_SHOULDER_RIGHT;
            break;
        case ButtonMapping::L2:
            report.Report.bTriggerL = 255;
            break;
        case ButtonMapping::R2:
            report.Report.bTriggerR = 255;
            break;
        case ButtonMapping::CROSS:
            report.Report.wButtons |= DS4_BUTTON_CROSS;
            break;
        case ButtonMapping::CIRCLE:
            report.Report.wButtons |= DS4_BUTTON_CIRCLE;
            break;
        case ButtonMapping::SQUARE:
            report.Report.wButtons |= DS4_BUTTON_SQUARE;
            break;
        case ButtonMapping::TRIANGLE:
            report.Report.wButtons |= DS4_BUTTON_TRIANGLE;
            break;
        case ButtonMapping::SHARE:
            report.Report.wButtons |= DS4_BUTTON_SHARE;
            break;
        case ButtonMapping::OPTIONS:
            report.Report.wButtons |= DS4_BUTTON_OPTIONS;
            break;
        case ButtonMapping::DPAD_UP:
            DS4_SET_DPAD(reinterpret_cast<PDS4_REPORT>(&report.Report), DS4_BUTTON_DPAD_NORTH);
            break;
        case ButtonMapping::DPAD_DOWN:
            DS4_SET_DPAD(reinterpret_cast<PDS4_REPORT>(&report.Report), DS4_BUTTON_DPAD_SOUTH);
            break;
        case ButtonMapping::DPAD_LEFT:
            DS4_SET_DPAD(reinterpret_cast<PDS4_REPORT>(&report.Report), DS4_BUTTON_DPAD_WEST);
            break;
        case ButtonMapping::DPAD_RIGHT:
            DS4_SET_DPAD(reinterpret_cast<PDS4_REPORT>(&report.Report), DS4_BUTTON_DPAD_EAST);
            break;
        case ButtonMapping::NONE:
        default:
            // Do nothing
            break;
    }
}

// Send keyboard key press/release using Windows SendInput API
void SendKeyboardInput(WORD virtualKey, bool keyDown) {
    INPUT input = {};
    input.type = INPUT_KEYBOARD;
    input.ki.wVk = virtualKey;
    input.ki.dwFlags = keyDown ? 0 : KEYEVENTF_KEYUP;
    input.ki.time = 0;
    input.ki.dwExtraInfo = 0;

    SendInput(1, &input, sizeof(INPUT));
}

// Track button states to avoid repeated key presses
static bool g_screenshotButtonPressed = false;
static bool g_cButtonPressed = false;
static bool g_comboPressed = false;

// Flag for ZL+ZR+GL+GR combo to enter management window
static std::atomic<bool> g_openManagementWindow(false);

// Handle special Pro Controller buttons (Screenshot, C button, and combo detection)
void HandleSpecialProButtons(const std::vector<uint8_t>& buffer) {
    if (buffer.size() < 9) return;

    // Build button state from bytes 3-8
    uint64_t state = 0;
    for (int i = 3; i <= 8; ++i) {
        state = (state << 8) | buffer[i];
    }

    constexpr uint64_t BUTTON_SCREENSHOT_MASK = 0x000020000000;  // Bit 29
    constexpr uint64_t BUTTON_C_MASK = 0x000040000000;           // Bit 30
    constexpr uint64_t BUTTON_GL_MASK = 0x000000000200;          // Bit 9
    constexpr uint64_t BUTTON_GR_MASK = 0x000000000100;          // Bit 8
    constexpr uint64_t TRIGGER_ZL_MASK = 0x000000800000;         // ZL trigger
    constexpr uint64_t TRIGGER_ZR_MASK = 0x008000000000;         // ZR trigger

    // Check for ZL+ZR+GL+GR combo to open management window (only trigger on initial press)
    bool comboCurrentlyPressed = (state & TRIGGER_ZL_MASK) && (state & TRIGGER_ZR_MASK) &&
                                  (state & BUTTON_GL_MASK) && (state & BUTTON_GR_MASK);

    if (comboCurrentlyPressed && !g_comboPressed) {
        // Combo just pressed - trigger management window
        g_openManagementWindow.store(true);
        g_comboPressed = true;
    }
    else if (!comboCurrentlyPressed && g_comboPressed) {
        // Combo released - reset state
        g_comboPressed = false;
    }

    // Handle Screenshot button -> F12 key
    bool screenshotPressed = (state & BUTTON_SCREENSHOT_MASK) != 0;
    if (screenshotPressed && !g_screenshotButtonPressed) {
        // Button just pressed - send F12 key down
        SendKeyboardInput(VK_F12, true);
        g_screenshotButtonPressed = true;
    }
    else if (!screenshotPressed && g_screenshotButtonPressed) {
        // Button just released - send F12 key up
        SendKeyboardInput(VK_F12, false);
        g_screenshotButtonPressed = false;
    }

    // Handle C button -> Cycle through layouts
    bool cPressed = (state & BUTTON_C_MASK) != 0;
    if (cPressed && !g_cButtonPressed) {
        // Button just pressed - cycle to next layout
        if (!g_proControllerConfig.layouts.empty()) {
            int oldIndex = g_proControllerConfig.activeLayoutIndex;
            g_proControllerConfig.activeLayoutIndex = (g_proControllerConfig.activeLayoutIndex + 1) % g_proControllerConfig.layouts.size();
            SaveProControllerConfig(g_proControllerConfig);

            // Log layout change
            std::cout << "\n[Layout Changed] "
                      << g_proControllerConfig.layouts[oldIndex].name
                      << " -> "
                      << g_proControllerConfig.layouts[g_proControllerConfig.activeLayoutIndex].name
                      << " (GL: " << ButtonMappingToString(g_proControllerConfig.layouts[g_proControllerConfig.activeLayoutIndex].glMapping)
                      << ", GR: " << ButtonMappingToString(g_proControllerConfig.layouts[g_proControllerConfig.activeLayoutIndex].grMapping)
                      << ")\n";
        }
        g_cButtonPressed = true;
    }
    else if (!cPressed && g_cButtonPressed) {
        g_cButtonPressed = false;
    }
}

// Apply GL/GR mappings to Pro Controller report using active layout
void ApplyGLGRMappings(DS4_REPORT_EX& report, const std::vector<uint8_t>& buffer) {
    if (buffer.size() < 9) return;

    // Check if we have any layouts
    if (g_proControllerConfig.layouts.empty()) return;

    // Get the active layout
    int layoutIndex = g_proControllerConfig.activeLayoutIndex;
    if (layoutIndex < 0 || layoutIndex >= static_cast<int>(g_proControllerConfig.layouts.size())) {
        layoutIndex = 0;
        g_proControllerConfig.activeLayoutIndex = 0;
    }

    const GLGRLayout& activeLayout = g_proControllerConfig.layouts[layoutIndex];

    // Build button state from bytes 3-8 (same as in JoyConDecoder)
    uint64_t state = 0;
    for (int i = 3; i <= 8; ++i) {
        state = (state << 8) | buffer[i];
    }

    constexpr uint64_t BUTTON_GL_MASK = 0x000000000200;  // Bit 9
    constexpr uint64_t BUTTON_GR_MASK = 0x000000000100;  // Bit 8

    // Apply mappings if buttons are pressed
    if (state & BUTTON_GL_MASK) {
        ApplyButtonMapping(report, activeLayout.glMapping);
    }
    if (state & BUTTON_GR_MASK) {
        ApplyButtonMapping(report, activeLayout.grMapping);
    }
}

void SendCustomCommands(GattCharacteristic const& characteristic)
{
    std::vector<std::vector<uint8_t>> commands = {
        { 0x0c, 0x91, 0x01, 0x02, 0x00, 0x04, 0x00, 0x00, 0xFF, 0x00, 0x00, 0x00 },
        { 0x0c, 0x91, 0x01, 0x04, 0x00, 0x04, 0x00, 0x00, 0xFF, 0x00, 0x00, 0x00 }
    };

    for (const auto& cmd : commands)
    {
        auto writer = DataWriter();
        writer.WriteBytes(cmd);
        IBuffer buffer = writer.DetachBuffer();

        auto status = characteristic.WriteValueAsync(buffer, GattWriteOption::WriteWithoutResponse).get();

        if (status == GattCommunicationStatus::Success)
        {
            std::wcout << L"Command sent successfully.\n";
        }
        else
        {
            std::wcout << L"Failed to send command.\n";
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(500));
    }
}
// Helper to send generic commands matching main.py structure
void SendGenericCommand(GattCharacteristic const& characteristic, uint8_t cmdId, uint8_t subCmdId, const std::vector<uint8_t>& data) {
    if (!characteristic) return;

    DataWriter writer;
    
    // Structure: CmdID, 0x91, 0x01, SubCmdID, 0x00, Len, 0x00, 0x00, Data...
    writer.WriteByte(cmdId);
    writer.WriteByte(0x91);
    writer.WriteByte(0x01);
    writer.WriteByte(subCmdId);
    writer.WriteByte(0x00);
    writer.WriteByte(static_cast<uint8_t>(data.size()));
    writer.WriteByte(0x00);
    writer.WriteByte(0x00);
    
    // Write Data
    for (uint8_t b : data) {
        writer.WriteByte(b);
    }

    IBuffer buffer = writer.DetachBuffer();
    characteristic.WriteValueAsync(buffer, GattWriteOption::WriteWithoutResponse).get();
    
    // Small delay to prevent flooding
    std::this_thread::sleep_for(std::chrono::milliseconds(50));
}

void EmitSound(GattCharacteristic const& characteristic) {
    // CMD 0x0A, SUB 0x02, Data: 0x04 (Preset) + Padding (up to 8 bytes total)
    std::vector<uint8_t> data(8, 0x00);
    data[0] = 0x04; // Preset ID
    SendGenericCommand(characteristic, 0x0A, 0x02, data);
}

void SetPlayerLEDs(GattCharacteristic const& characteristic, uint8_t pattern) {
    // CMD 0x09, SUB 0x07, Data: pattern + Padding (up to 8 bytes total)
    std::vector<uint8_t> data(8, 0x00);
    data[0] = pattern;
    SendGenericCommand(characteristic, 0x09, 0x07, data);
}

struct ConnectedJoyCon {
    BluetoothLEDevice device = nullptr;
    GattCharacteristic inputChar = nullptr;
    GattCharacteristic writeChar = nullptr;
};

const wchar_t* GattStatusToString(GattCommunicationStatus status)
{
    switch (status) {
        case GattCommunicationStatus::Success:
            return L"Success";
        case GattCommunicationStatus::Unreachable:
            return L"Unreachable";
        case GattCommunicationStatus::ProtocolError:
            return L"ProtocolError";
        case GattCommunicationStatus::AccessDenied:
            return L"AccessDenied";
        default:
            return L"Unknown";
    }
}

ConnectedJoyCon WaitForJoyCon(const std::wstring& prompt)
{
    std::wcout << prompt << L"\n";

    ConnectedJoyCon cj{};

    BluetoothLEDevice device = nullptr;
    bool connected = false;

    BluetoothLEAdvertisementWatcher watcher;

    std::mutex mtx;
    std::condition_variable cv;

    watcher.Received([&](auto const&, auto const& args)
        {
            std::unique_lock<std::mutex> lock(mtx);
            if (connected) return;

            auto mfg = args.Advertisement().ManufacturerData();
            for (uint32_t i = 0; i < mfg.Size(); i++)
            {
                auto section = mfg.GetAt(i);
                if (section.CompanyId() != JOYCON_MANUFACTURER_ID) continue;
                auto reader = DataReader::FromBuffer(section.Data());
                std::vector<uint8_t> data(reader.UnconsumedBufferLength());
                reader.ReadBytes(data);
                if (data.size() >= JOYCON_MANUFACTURER_PREFIX.size() &&
                    std::equal(JOYCON_MANUFACTURER_PREFIX.begin(), JOYCON_MANUFACTURER_PREFIX.end(), data.begin()))
                {
                    device = BluetoothLEDevice::FromBluetoothAddressAsync(args.BluetoothAddress()).get();
                    if (!device) return;

                    connected = true;
                    watcher.Stop();
                    cv.notify_one();
                    return;
                }
            }
        });

    watcher.ScanningMode(BluetoothLEScanningMode::Active);
    watcher.Start();

    std::wcout << L"Scanning for Joy-Con... (Waiting up to 30 seconds)\n";

    {
        std::unique_lock<std::mutex> lock(mtx);
        if (!cv.wait_for(lock, std::chrono::seconds(30), [&]() { return connected; }))
        {
            watcher.Stop();
            std::wcerr << L"Timeout: Joy-Con not found.\n";
            exit(1);
        }
    }

    cj.device = device;

    GattDeviceServicesResult servicesResult = nullptr;
    for (int attempt = 1; attempt <= 10; ++attempt) {
        servicesResult = device.GetGattServicesAsync(BluetoothCacheMode::Uncached).get();
        if (servicesResult.Status() == GattCommunicationStatus::Success) {
            break;
        }

        std::wcerr << L"Failed to get GATT services (attempt " << attempt
            << L"/10, status=" << GattStatusToString(servicesResult.Status()) << L"). Retrying...\n";
        std::this_thread::sleep_for(std::chrono::milliseconds(500));
    }

    if (!servicesResult || servicesResult.Status() != GattCommunicationStatus::Success) {
        std::wcerr << L"Failed to get GATT services. Remove the Joy-Con from Windows Bluetooth settings, turn it off, then pair it again.\n";
        exit(1);
    }

    for (auto service : servicesResult.Services())
    {
        auto charsResult = service.GetCharacteristicsAsync().get();
        if (charsResult.Status() != GattCommunicationStatus::Success) continue;
        for (auto characteristic : charsResult.Characteristics())
        {
            if (characteristic.Uuid() == guid(INPUT_REPORT_UUID))
                cj.inputChar = characteristic;
            else if (characteristic.Uuid() == guid(WRITE_COMMAND_UUID))
                cj.writeChar = characteristic;
        }
    }

    if (!cj.inputChar || !cj.writeChar) {
        std::wcerr << L"Joy-Con connected, but required GATT characteristics were not found.\n";
        std::wcerr << L"Try removing the device from Windows Bluetooth settings and pairing it again.\n";
        exit(1);
    }

    return cj;
}

enum ControllerType {
    SingleJoyCon = 1,
    DualJoyCon = 2,
    ProController = 3,
    NSOGCController = 4
};

struct PlayerConfig {
    ControllerType controllerType;
    JoyConSide joyconSide;
    JoyConOrientation joyconOrientation;
    GyroSource gyroSource;
    GyroMode gyroMode;
    MotionProfile motionProfile;
};

// For single Joy-Con players, store controller + JoyCon info to keep alive
struct SingleJoyConPlayer {
    ConnectedJoyCon joycon;
    PVIGEM_TARGET ds4Controller;
    JoyConSide side;
    JoyConOrientation orientation;
    
    // Mouse State
    int mouseMode = 0; // 0=Off, 1=Fast, 2=Normal, 3=Slow
    bool wasChatPressed = false;
    int16_t lastOpticalX = 0;
    int16_t lastOpticalY = 0;
    bool firstOpticalRead = true;
    
    // Scroll Accumulator
    float scrollAccumulator = 0.0f;
    
    // Button States for Edge Detection (to avoid rapid fire)
    bool mb4Pressed = false;
    bool mb5Pressed = false;
    
    // Previous button states for click emulation
    bool leftBtnPressed = false;
    bool rightBtnPressed = false;
    bool middleBtnPressed = false;

    LatencyTracker latency;
};

struct TimedInputBuffer {
    std::vector<uint8_t> buffer;
    TimePoint receivedAt{};
    double bleDeltaMs = -1.0;
    uint64_t sequence = 0;
};

struct DualJoyConSharedState {
    std::mutex mutex;
    std::condition_variable cv;
    TimedInputBuffer left;
    TimedInputBuffer right;
    TimePoint lastLeftBleTime{};
    TimePoint lastRightBleTime{};
    TimePoint lastEmitTime{};
    uint64_t sequence = 0;
    uint64_t eventIndex = 0;
};

// For dual Joy-Con players, store both JoyCons, controller, thread, and running flag
struct DualJoyConPlayer {
    ConnectedJoyCon leftJoyCon;
    ConnectedJoyCon rightJoyCon;
    GyroSource gyroSource;
    GyroMode gyroMode;
    MotionProfile motionProfile;
    uint8_t dsuSlot = 0;
    PVIGEM_TARGET ds4Controller;
    std::atomic<bool> running;
    std::thread updateThread;
    std::shared_ptr<DualJoyConSharedState> sharedState;
};

// For Pro Controller players
struct ProControllerPlayer {
    ConnectedJoyCon controller;
    PVIGEM_TARGET ds4Controller;
    LatencyTracker latency;
};

int main(int argc, char* argv[])
{
    init_apartment();

    for (int arg = 1; arg < argc; ++arg) {
        const std::string option = argv[arg];
        if (option == "--latency-test" || option == "--latency-benchmark") {
            g_runtimeOptions.latencyMetrics = true;
        }
        else if (option == "--update-policy" && arg + 1 < argc) {
            const std::string value = argv[++arg];
            if (value == "low" || value == "LowLatency") {
                g_runtimeOptions.updatePolicy = UpdatePolicy::LowLatency;
            }
            else if (value == "balanced" || value == "Balanced120Hz") {
                g_runtimeOptions.updatePolicy = UpdatePolicy::Balanced120Hz;
            }
            else if (value == "legacy" || value == "Legacy60Hz") {
                g_runtimeOptions.updatePolicy = UpdatePolicy::Legacy60Hz;
            }
        }
        else if (option == "--latency-csv" && arg + 1 < argc) {
            g_runtimeOptions.latencyCsvPath = argv[++arg];
        }
    }

    std::wstring metricsLine;
    std::wcout << L"Enable latency metrics? (y/N): ";
    std::getline(std::wcin, metricsLine);
    if (metricsLine == L"y" || metricsLine == L"Y") {
        g_runtimeOptions.latencyMetrics = true;
    }

    std::wstring policyLine;
    std::wcout << L"Update policy? (1=LowLatency, 2=Balanced120Hz, 3=Legacy60Hz) [1]: ";
    std::getline(std::wcin, policyLine);
    if (policyLine == L"2") {
        g_runtimeOptions.updatePolicy = UpdatePolicy::Balanced120Hz;
    }
    else if (policyLine == L"3") {
        g_runtimeOptions.updatePolicy = UpdatePolicy::Legacy60Hz;
    }

    if (g_runtimeOptions.latencyMetrics) {
        if (g_latencyLogger.Start(g_runtimeOptions.latencyCsvPath)) {
            const std::wstring csvPath(g_runtimeOptions.latencyCsvPath.begin(), g_runtimeOptions.latencyCsvPath.end());
            std::wcout << L"Latency metrics enabled: " << csvPath << L"\n";
        }
        else {
            std::wcerr << L"Warning: could not open latency CSV. Metrics disabled.\n";
        }
    }
    const std::string policyName = UpdatePolicyName(g_runtimeOptions.updatePolicy);
    std::wcout << L"Update policy: " << std::wstring(policyName.begin(), policyName.end()) << L"\n";

    int numPlayers;
    std::wcout << L"How many players? ";
    std::wcin >> numPlayers;
    std::wcin.ignore();

    std::vector<PlayerConfig> playerConfigs;

    for (int i = 0; i < numPlayers; ++i) {
        PlayerConfig config{};
        std::wstring line;

        while (true) {
            std::wcout << L"Player " << (i + 1) << L":\n";
            std::wcout << L"  What controller type? (1=Single JoyCon, 2=Dual JoyCon, 3=Pro Controller, 4=NSO GC Controller): ";
            std::getline(std::wcin, line);
            if (line == L"1" || line == L"2" || line == L"3" || line == L"4") {
                config.controllerType = static_cast<ControllerType>(std::stoi(std::string(line.begin(), line.end())));
                break;
            }
            std::wcout << L"Invalid input. Please enter 1, 2, or 3.\n";
        }

        if (config.controllerType == SingleJoyCon) {
            while (true) {
                std::wcout << L"  Which side? (L=Left, R=Right): ";
                std::getline(std::wcin, line);
                if (line == L"L" || line == L"R" || line == L"l" || line == L"r") {
                    config.joyconSide = (line == L"L" || line == L"l") ? JoyConSide::Left : JoyConSide::Right;
                    break;
                }
                std::wcout << L"Invalid input. Please enter L or R.\n";
            }
            while (true) {
                std::wcout << L"  What orientation? (U=Upright, S=Sideways): ";
                std::getline(std::wcin, line);
                if (line == L"U" || line == L"S" || line == L"u" || line == L"s") {
                    config.joyconOrientation = (line == L"S" || line == L"s") ? JoyConOrientation::Sideways : JoyConOrientation::Upright;
                    break;
                }
                std::wcout << L"Invalid input. Please enter U or S.\n";
            }
        }
        else if (config.controllerType == DualJoyCon) {
            config.joyconSide = JoyConSide::Left;
            config.joyconOrientation = JoyConOrientation::Upright;

            while (true) {
                std::wcout << L"Player " << (i + 1) << L":\n";
                std::wcout << L"  What should be used as Gyro Source? (B=Both JoyCons, L=Left JoyCon, R=Right JoyCon): ";
                std::getline(std::wcin, line);

                if (line == L"B") {
                    config.gyroSource = GyroSource::Both;
                    break;
                }
                else if (line == L"L") {
                    config.gyroSource = GyroSource::Left;
                    break;
                }
                else if (line == L"R") {
                    config.gyroSource = GyroSource::Right;
                    break;
                }
                std::wcout << L"Invalid input. Please enter B, L, or R.\n";
            }

        }

        if (config.controllerType == SingleJoyCon || config.controllerType == DualJoyCon || config.controllerType == ProController) {
            while (true) {
                std::wcout << L"  Gyro output? (1=DS4 Raw, 2=DS4 Switch Emulator, 3=DSU UDP): ";
                std::getline(std::wcin, line);
                if (line == L"1" || line.empty()) {
                    config.gyroMode = GyroMode::DS4Raw;
                    config.motionProfile = MotionProfile::Raw;
                    break;
                }
                if (line == L"2") {
                    config.gyroMode = GyroMode::DS4SwitchEmu;
                    config.motionProfile = MotionProfile::SwitchEmu;
                    break;
                }
                if (line == L"3") {
                    config.gyroMode = GyroMode::DsuUdp;
                    config.motionProfile = MotionProfile::SwitchEmu;
                    std::wcout << L"  Configure your emulator's Cemuhook/DSU motion source at 127.0.0.1:26760.\n";
                    if (config.controllerType == DualJoyCon && config.gyroSource == GyroSource::Both) {
                        std::wcout << L"  Tip: choose Right as gyro source for pointer/sword motion when using the right Joy-Con.\n";
                    }
                    break;
                }
                std::wcout << L"Invalid input. Please enter 1, 2, or 3.\n";
            }
        }
        else {
            config.gyroMode = GyroMode::DS4Raw;
            config.motionProfile = MotionProfile::Raw;
        }

        playerConfigs.push_back(config);
    }

    DsuServer dsuServer;
    const bool needsDsu = std::any_of(playerConfigs.begin(), playerConfigs.end(), [](const PlayerConfig& config) {
        return config.gyroMode == GyroMode::DsuUdp;
    });
    if (needsDsu) {
        if (dsuServer.Start()) {
            std::wcout << L"DSU UDP server listening on 127.0.0.1:26760.\n";
        }
        else {
            std::wcerr << L"Failed to start DSU UDP server on 127.0.0.1:26760.\n";
        }
    }

    InitializeViGEm();

    // Store all players to keep them alive
    std::vector<SingleJoyConPlayer> singlePlayers;
    std::vector<std::unique_ptr<DualJoyConPlayer>> dualPlayers;
    std::vector<ProControllerPlayer> proPlayers;
    singlePlayers.reserve(numPlayers);
    dualPlayers.reserve(numPlayers);
    proPlayers.reserve(numPlayers);

    for (int i = 0; i < numPlayers; ++i) {
        auto& config = playerConfigs[i];
        std::wcout << L"Player " << (i + 1) << L" setup...\n";

        if (config.controllerType == SingleJoyCon) {
            std::wstring sideStr = (config.joyconSide == JoyConSide::Left) ? L"Left" : L"Right";
            std::wcout << L"Please sync your single Joy-Con (" << sideStr << L") now.\n";

            ConnectedJoyCon cj = WaitForJoyCon(L"Waiting for single Joy-Con...");

            // Request minimum BLE connection interval for lowest latency
            try {
                auto connectionParams = BluetoothLEPreferredConnectionParameters::ThroughputOptimized();
                cj.device.RequestPreferredConnectionParameters(connectionParams);
                std::wcout << L"Requested ThroughputOptimized connection parameters for lower latency (Windows may ignore this request).\n";
            }
            catch (...) {
                std::wcout << L"Warning: Could not request preferred connection parameters.\n";
            }

            PVIGEM_TARGET ds4_controller = vigem_target_ds4_alloc();
            auto ret = vigem_target_add(vigem_client, ds4_controller);
            if (!VIGEM_SUCCESS(ret))
            {
                std::wcerr << L"Failed to add DS4 controller target: 0x" << std::hex << ret << L"\n";
                exit(1);
            }

            singlePlayers.push_back({ cj, ds4_controller, config.joyconSide, config.joyconOrientation });
            auto& player = singlePlayers.back();

            if (config.gyroMode == GyroMode::DsuUdp && dsuServer.IsRunning()) {
                dsuServer.SetControllerConnected(static_cast<uint8_t>(i));
            }

            player.joycon.inputChar.ValueChanged([joyconSide = player.side, joyconOrientation = player.orientation, &player, motionProfile = config.motionProfile, gyroMode = config.gyroMode, dsuSlot = static_cast<uint8_t>(i), &dsuServer](GattCharacteristic const&, GattValueChangedEventArgs const& args)
                {
                    const auto notificationReceived = SteadyClock::now();
                    auto reader = DataReader::FromBuffer(args.CharacteristicValue());
                    std::vector<uint8_t> buffer(reader.UnconsumedBufferLength());
                    reader.ReadBytes(buffer);
                    const double bleDeltaMs = MillisecondsBetween(player.latency.lastBleTime, notificationReceived);
                    player.latency.lastBleTime = notificationReceived;

                    // Optical Mouse Toggle Logic (Only for Right Joy-Con/Joy-Con 2)
                    if (joyconSide == JoyConSide::Right) {
                        uint32_t btnState = ExtractButtonState(buffer);
                        // CHAT button mask 0x000040 (Right JoyCon)
                        bool chatPressed = (btnState & 0x000040) != 0;

                        if (chatPressed && !player.wasChatPressed) {
                            player.mouseMode = (player.mouseMode + 1) % 4;
                            const char* modeName = "OFF";
                            uint8_t ledPattern = 0x01; // Default OFF = LED 1
                            if (player.mouseMode == 1) {
                                modeName = "FAST";
                                ledPattern = 0x02; // LED 2
                            }
                            else if (player.mouseMode == 2) {
                                modeName = "NORMAL";
                                ledPattern = 0x04; // LED 3
                            }
                            else if (player.mouseMode == 3) {
                                modeName = "SLOW";
                                ledPattern = 0x08; // LED 4
                            }
                            
                            std::cout << "Optical Mouse Mode: " << modeName << std::endl;
                            SetPlayerLEDs(player.joycon.writeChar, ledPattern);
                            EmitSound(player.joycon.writeChar);
                        }
                        player.wasChatPressed = chatPressed;
                        
                        // If Mouse Mode is ON (1, 2, or 3)
                        if (player.mouseMode > 0) {
                            // --- 1. Optical Mouse Movement ---
                            auto [rawX, rawY] = GetRawOpticalMouse(buffer);
                            if (player.firstOpticalRead) {
                                player.lastOpticalX = rawX;
                                player.lastOpticalY = rawY;
                                player.firstOpticalRead = false;
                            } else {
                                int16_t dx = rawX - player.lastOpticalX;
                                int16_t dy = rawY - player.lastOpticalY;
                                player.lastOpticalX = rawX;
                                player.lastOpticalY = rawY;

                                if (dx != 0 || dy != 0) {
                                    float sensitivity = 1.0f;
                                    if (player.mouseMode == 1) sensitivity = 1.0f;      // TODO: Fast 可调
                                    else if (player.mouseMode == 2) sensitivity = 0.6f; // TODO: Normal 可调
                                    else if (player.mouseMode == 3) sensitivity = 0.3f; // TODO: Slow 可调
                                    
                                    int moveX = static_cast<int>(dx * sensitivity);
                                    int moveY = static_cast<int>(dy * sensitivity);
                                    
                                    INPUT input = {};
                                    input.type = INPUT_MOUSE;
                                    input.mi.dx = moveX;
                                    input.mi.dy = moveY;
                                    input.mi.dwFlags = MOUSEEVENTF_MOVE;
                                    SendInput(1, &input, sizeof(INPUT));
                                }
                            }
                            
                            // --- 2. Mouse Buttons (R=Left, ZR=Right, Stick=Middle) ---
                            // R = 0x004000, ZR = 0x008000, Stick = 0x000004
                            bool rPressed = (btnState & 0x004000) != 0;
                            bool zrPressed = (btnState & 0x008000) != 0;
                            bool stickPressed = (btnState & 0x000004) != 0;
                            
                            // Left Click (R)
                            if (rPressed && !player.leftBtnPressed) {
                                INPUT input = {}; input.type = INPUT_MOUSE; input.mi.dwFlags = MOUSEEVENTF_LEFTDOWN; SendInput(1, &input, sizeof(INPUT));
                            } else if (!rPressed && player.leftBtnPressed) {
                                INPUT input = {}; input.type = INPUT_MOUSE; input.mi.dwFlags = MOUSEEVENTF_LEFTUP; SendInput(1, &input, sizeof(INPUT));
                            }
                            player.leftBtnPressed = rPressed;
                            
                            // Right Click (ZR)
                            if (zrPressed && !player.rightBtnPressed) {
                                INPUT input = {}; input.type = INPUT_MOUSE; input.mi.dwFlags = MOUSEEVENTF_RIGHTDOWN; SendInput(1, &input, sizeof(INPUT));
                            } else if (!zrPressed && player.rightBtnPressed) {
                                INPUT input = {}; input.type = INPUT_MOUSE; input.mi.dwFlags = MOUSEEVENTF_RIGHTUP; SendInput(1, &input, sizeof(INPUT));
                            }
                            player.rightBtnPressed = zrPressed;
                            
                            // Middle Click (Stick)
                            if (stickPressed && !player.middleBtnPressed) {
                                INPUT input = {}; input.type = INPUT_MOUSE; input.mi.dwFlags = MOUSEEVENTF_MIDDLEDOWN; SendInput(1, &input, sizeof(INPUT));
                            } else if (!stickPressed && player.middleBtnPressed) {
                                INPUT input = {}; input.type = INPUT_MOUSE; input.mi.dwFlags = MOUSEEVENTF_MIDDLEUP; SendInput(1, &input, sizeof(INPUT));
                            }
                            player.middleBtnPressed = stickPressed;
                            
                            // --- 3. Stick Scrolling & Side Buttons ---
                            auto stickData = DecodeJoystick(buffer, joyconSide, joyconOrientation);
                            
                            // Scroll (Vertical Stick)
                            // stickData.y is normalized int16 (-32767 to 32767).
                            // Up is typically negative Y in raw data?, DecodeJoystick returns standard cartesian?
                            // Let's check testapp logic: outY = -y * 32767. If y was (y_raw - 2048), then raw 4095 (down) -> +y -> outY negative.
                            // So DecodeJoystick: Up (+y_raw?) -> outY Positive?
                            // Let's assume standard behavior: Up is Positive Y, Down is Negative Y in stickData (based on math).
                            // Windows Wheel: +Delta is Up, -Delta is Down.
                            
                            // Let's try: Stick Up -> Scroll Up.
                            // If stickData.y > deadzone -> Scroll Up.
                            // Speed proportional to magnitude.
                            
                            const int SCROLL_DEADZONE = 4000;
                            if (abs(stickData.y) > SCROLL_DEADZONE) {
                                // Accumulate scroll
                                float intensity = (abs(stickData.y) - SCROLL_DEADZONE) / (32767.0f - SCROLL_DEADZONE);
                                float speed = intensity * 40.0f; // Max scroll per frame // TODO: 滚轮速度可调
                                if (stickData.y > 0) player.scrollAccumulator -= speed; // Up
                                else player.scrollAccumulator += speed; // Down
                                
                                if (abs(player.scrollAccumulator) >= 120.0f) {
                                    int clicks = static_cast<int>(player.scrollAccumulator / 120.0f);
                                    player.scrollAccumulator -= (clicks * 120.0f);
                                    
                                    INPUT input = {};
                                    input.type = INPUT_MOUSE;
                                    input.mi.mouseData = clicks * 120;
                                    input.mi.dwFlags = MOUSEEVENTF_WHEEL;
                                    SendInput(1, &input, sizeof(INPUT));
                                }
                            } else {
                                player.scrollAccumulator = 0.0f;
                            }
                            
                            // Side Buttons (Horizontal Stick)
                            // Left Peak -> Back (MB4)
                            // Right Peak -> Forward (MB5)
                            const int BUTTON_THRESHOLD = 28000; // Near edge
                            
                            // MB4 (Back) - Left
                            if (stickData.x < -BUTTON_THRESHOLD) {
                                if (!player.mb4Pressed) {
                                    INPUT input = {}; input.type = INPUT_MOUSE; input.mi.mouseData = XBUTTON1; input.mi.dwFlags = MOUSEEVENTF_XDOWN; SendInput(1, &input, sizeof(INPUT));
                                    INPUT input2 = {}; input2.type = INPUT_MOUSE; input2.mi.mouseData = XBUTTON1; input2.mi.dwFlags = MOUSEEVENTF_XUP; SendInput(1, &input2, sizeof(INPUT));
                                    player.mb4Pressed = true;
                                }
                            } else {
                                player.mb4Pressed = false;
                            }
                            
                            // MB5 (Forward) - Right
                            if (stickData.x > BUTTON_THRESHOLD) {
                                if (!player.mb5Pressed) {
                                    INPUT input = {}; input.type = INPUT_MOUSE; input.mi.mouseData = XBUTTON2; input.mi.dwFlags = MOUSEEVENTF_XDOWN; SendInput(1, &input, sizeof(INPUT));
                                    INPUT input2 = {}; input2.type = INPUT_MOUSE; input2.mi.mouseData = XBUTTON2; input2.mi.dwFlags = MOUSEEVENTF_XUP; SendInput(1, &input2, sizeof(INPUT));
                                    player.mb5Pressed = true;
                                }
                            } else {
                                player.mb5Pressed = false;
                            }
                            
                            // --- 4. Suppress Inputs in DS4 Report ---
                            // Modify buffer to clear mapped buttons so they don't trigger game actions
                            // R (Byte 3, bit 6: 0x40), ZR (Byte 4, bit 7: 0x80 ... wait, let's check masks)
                            
                            // Button masks again:
                            // R: 0x004000 -> Byte 3 & 0x40.
                            // ZR: 0x008000 -> Byte 4 & 0x80 (Wait, 0x008000 is 3rd byte of 24-bit? 3,4,5. 0x004000 is 0x40 << 8? No)
                            // ExtractButtonState: (buffer[3] << 16) | (buffer[4] << 8) | buffer[5]
                            // 0x004000 = Bit 14 set.
                            // buffer[3] is bits 23-16. buffer[4] is 15-8. buffer[5] is 7-0.
                            // 0x4000 is in buffer[4] (0x40 << 8).
                            // 0x8000 is in buffer[4] (0x80 << 8).
                            // Stick Click 0x04 is in buffer[5].
                            
                            // So:
                            // Clear R bit (0x40 in buffer[4])
                            buffer[4] &= ~0x40;
                            // Clear ZR bit (0x80 in buffer[4])
                            buffer[4] &= ~0x80;
                            // Clear Stick Click (0x04 in buffer[5])
                            buffer[5] &= ~0x04;
                            
                            // Clear Stick Data (Set to center)
                            // Right stick bytes: 13, 14, 15
                            if (buffer.size() >= 16) {
                                buffer[13] = 0x00;
                                buffer[14] = 0x08;
                                buffer[15] = 0x80;
                            }
                        } else {
                            // Mode Off: Reset first reads
                            player.firstOpticalRead = true;
                        }
                    }

                    if (!ShouldEmitForPolicy(g_runtimeOptions.updatePolicy, player.latency.lastEmitTime, SteadyClock::now())) {
                        return;
                    }

                    const auto decodeStart = SteadyClock::now();
                    const MotionProfile ds4Profile = (gyroMode == GyroMode::DsuUdp) ? MotionProfile::Raw : motionProfile;
                    DS4_REPORT_EX report = GenerateDS4Report(buffer, joyconSide, joyconOrientation, ds4Profile);

                    if (gyroMode == GyroMode::DsuUdp && dsuServer.IsRunning()) {
                        DS4_REPORT_EX dsuReport = GenerateDS4Report(buffer, joyconSide, joyconOrientation, MotionProfile::SwitchEmu);
                        dsuServer.UpdateController(dsuSlot, dsuReport);
                    }

                    auto ret = vigem_target_ds4_update_ex(vigem_client, player.ds4Controller, report);
                    const auto vigemComplete = SteadyClock::now();
                    if (!VIGEM_SUCCESS(ret)) {
                         // std::wcerr << L"Failed to update DS4 EX report: 0x" << std::hex << ret << L"\n";
                    }

                    g_latencyLogger.Record(
                        g_runtimeOptions.updatePolicy,
                        ControllerTypeName(SingleJoyCon),
                        ++player.latency.eventIndex,
                        bleDeltaMs,
                        0.0,
                        -1.0,
                        MicrosecondsBetween(decodeStart, vigemComplete),
                        MicrosecondsBetween(notificationReceived, vigemComplete));
                });

            auto status = player.joycon.inputChar.WriteClientCharacteristicConfigurationDescriptorAsync(
                GattClientCharacteristicConfigurationDescriptorValue::Notify).get();

            if (player.joycon.writeChar) {
                SendCustomCommands(player.joycon.writeChar);
                std::this_thread::sleep_for(std::chrono::milliseconds(200));
                SetPlayerLEDs(player.joycon.writeChar, 0x01); // Player 1 (Solid LED 1)
                EmitSound(player.joycon.writeChar);
            }

            if (status == GattCommunicationStatus::Success)
                std::wcout << L"Notifications enabled.\n";
            else
                std::wcout << L"Failed to enable notifications.\n";

            std::wcout << L"Press Enter to continue...\n";
            std::wstring dummy;
            std::getline(std::wcin, dummy);
        }
        else if (config.controllerType == DualJoyCon) {
            std::wcout << L"Please sync your RIGHT Joy-Con now.\n";
            ConnectedJoyCon rightJoyCon = WaitForJoyCon(L"Waiting for RIGHT Joy-Con...");
            if (rightJoyCon.writeChar) {
                SendCustomCommands(rightJoyCon.writeChar);
                std::this_thread::sleep_for(std::chrono::milliseconds(200));
                SetPlayerLEDs(rightJoyCon.writeChar, 0x01);
                EmitSound(rightJoyCon.writeChar);
            }

            // Request minimum BLE connection interval for right Joy-Con
            try {
                auto connectionParams = BluetoothLEPreferredConnectionParameters::ThroughputOptimized();
                rightJoyCon.device.RequestPreferredConnectionParameters(connectionParams);
                std::wcout << L"Requested ThroughputOptimized connection parameters for RIGHT Joy-Con (Windows may ignore this request).\n";
            }
            catch (...) {
                std::wcout << L"Warning: Could not request preferred connection parameters for RIGHT Joy-Con.\n";
            }

            std::wcout << L"Please sync your LEFT Joy-Con now.\n";
            ConnectedJoyCon leftJoyCon = WaitForJoyCon(L"Waiting for LEFT Joy-Con...");
            if (leftJoyCon.writeChar) {
                SendCustomCommands(leftJoyCon.writeChar);
                std::this_thread::sleep_for(std::chrono::milliseconds(200));
                SetPlayerLEDs(leftJoyCon.writeChar, 0x08); // Player 2 (Solid LED 4? or specific pattern for left?)
                EmitSound(leftJoyCon.writeChar);
            }

            // Request minimum BLE connection interval for left Joy-Con
            try {
                auto connectionParams = BluetoothLEPreferredConnectionParameters::ThroughputOptimized();
                leftJoyCon.device.RequestPreferredConnectionParameters(connectionParams);
                std::wcout << L"Requested ThroughputOptimized connection parameters for LEFT Joy-Con (Windows may ignore this request).\n";
            }
            catch (...) {
                std::wcout << L"Warning: Could not request preferred connection parameters for LEFT Joy-Con.\n";
            }

            PVIGEM_TARGET ds4Controller = vigem_target_ds4_alloc();
            auto ret = vigem_target_add(vigem_client, ds4Controller);
            if (!VIGEM_SUCCESS(ret))
            {
                std::wcerr << L"Failed to add DS4 controller target: 0x" << std::hex << ret << L"\n";
                exit(1);
            }

            auto dualPlayer = std::make_unique<DualJoyConPlayer>();
            dualPlayer->leftJoyCon = leftJoyCon;
            dualPlayer->rightJoyCon = rightJoyCon;
            dualPlayer->gyroSource = config.gyroSource;
            dualPlayer->gyroMode = config.gyroMode;
            dualPlayer->motionProfile = config.motionProfile;
            dualPlayer->dsuSlot = static_cast<uint8_t>(i);
            dualPlayer->ds4Controller = ds4Controller;
            dualPlayer->running.store(true);
            dualPlayer->sharedState = std::make_shared<DualJoyConSharedState>();

            if (dualPlayer->gyroMode == GyroMode::DsuUdp && dsuServer.IsRunning()) {
                dsuServer.SetControllerConnected(dualPlayer->dsuSlot);
            }

            auto dualSharedState = dualPlayer->sharedState;

            dualPlayer->leftJoyCon.inputChar.ValueChanged([dualSharedState](GattCharacteristic const&, GattValueChangedEventArgs const& args)
                {
                    const auto receivedAt = SteadyClock::now();
                    auto reader = DataReader::FromBuffer(args.CharacteristicValue());
                    std::vector<uint8_t> buffer(reader.UnconsumedBufferLength());
                    reader.ReadBytes(buffer);

                    {
                        std::lock_guard<std::mutex> lock(dualSharedState->mutex);
                        dualSharedState->left.bleDeltaMs = MillisecondsBetween(dualSharedState->lastLeftBleTime, receivedAt);
                        dualSharedState->lastLeftBleTime = receivedAt;
                        dualSharedState->left.buffer = std::move(buffer);
                        dualSharedState->left.receivedAt = receivedAt;
                        dualSharedState->left.sequence = ++dualSharedState->sequence;
                    }
                    dualSharedState->cv.notify_one();
                });

            auto statusLeft = dualPlayer->leftJoyCon.inputChar.WriteClientCharacteristicConfigurationDescriptorAsync(
                GattClientCharacteristicConfigurationDescriptorValue::Notify).get();

            if (statusLeft == GattCommunicationStatus::Success)
                std::wcout << L"LEFT Joy-Con notifications enabled.\n";
            else
                std::wcout << L"Failed to enable LEFT Joy-Con notifications.\n";

            dualPlayer->rightJoyCon.inputChar.ValueChanged([dualSharedState](GattCharacteristic const&, GattValueChangedEventArgs const& args)
                {
                    const auto receivedAt = SteadyClock::now();
                    auto reader = DataReader::FromBuffer(args.CharacteristicValue());
                    std::vector<uint8_t> buffer(reader.UnconsumedBufferLength());
                    reader.ReadBytes(buffer);

                    {
                        std::lock_guard<std::mutex> lock(dualSharedState->mutex);
                        dualSharedState->right.bleDeltaMs = MillisecondsBetween(dualSharedState->lastRightBleTime, receivedAt);
                        dualSharedState->lastRightBleTime = receivedAt;
                        dualSharedState->right.buffer = std::move(buffer);
                        dualSharedState->right.receivedAt = receivedAt;
                        dualSharedState->right.sequence = ++dualSharedState->sequence;
                    }
                    dualSharedState->cv.notify_one();
                });

            auto statusRight = dualPlayer->rightJoyCon.inputChar.WriteClientCharacteristicConfigurationDescriptorAsync(
                GattClientCharacteristicConfigurationDescriptorValue::Notify).get();

            if (statusRight == GattCommunicationStatus::Success)
                std::wcout << L"RIGHT Joy-Con notifications enabled.\n";
            else
                std::wcout << L"Failed to enable RIGHT Joy-Con notifications.\n";

            dualPlayer->updateThread = std::thread([dualPlayerPtr = dualPlayer.get(), dualSharedState, &dsuServer]()
                {
                    uint64_t lastProcessedSequence = 0;
                    while (dualPlayerPtr->running.load(std::memory_order_acquire))
                    {
                        TimedInputBuffer leftSnapshot;
                        TimedInputBuffer rightSnapshot;

                        {
                            std::unique_lock<std::mutex> lock(dualSharedState->mutex);
                            if (g_runtimeOptions.updatePolicy == UpdatePolicy::Legacy60Hz) {
                                dualSharedState->cv.wait_for(lock, std::chrono::milliseconds(16), [&]() {
                                    return !dualPlayerPtr->running.load(std::memory_order_acquire) ||
                                           dualSharedState->sequence != lastProcessedSequence;
                                });
                            }
                            else {
                                dualSharedState->cv.wait_for(lock, std::chrono::milliseconds(1), [&]() {
                                    return !dualPlayerPtr->running.load(std::memory_order_acquire) ||
                                           dualSharedState->sequence != lastProcessedSequence;
                                });
                            }

                            if (!dualPlayerPtr->running.load(std::memory_order_acquire)) {
                                break;
                            }

                            if (dualSharedState->left.buffer.empty() || dualSharedState->right.buffer.empty() ||
                                dualSharedState->sequence == lastProcessedSequence) {
                                continue;
                            }

                            const auto now = SteadyClock::now();
                            if (!ShouldEmitForPolicy(g_runtimeOptions.updatePolicy, dualSharedState->lastEmitTime, now)) {
                                lastProcessedSequence = dualSharedState->sequence;
                                continue;
                            }

                            leftSnapshot = dualSharedState->left;
                            rightSnapshot = dualSharedState->right;
                            lastProcessedSequence = dualSharedState->sequence;
                        }

                        const auto decodeStart = SteadyClock::now();
                        const MotionProfile ds4Profile = (dualPlayerPtr->gyroMode == GyroMode::DsuUdp) ? MotionProfile::Raw : dualPlayerPtr->motionProfile;
                        DS4_REPORT_EX report = GenerateDualJoyConDS4Report(leftSnapshot.buffer, rightSnapshot.buffer, dualPlayerPtr->gyroSource, ds4Profile);

                        if (dualPlayerPtr->gyroMode == GyroMode::DsuUdp && dsuServer.IsRunning()) {
                            DS4_REPORT_EX dsuReport = GenerateDualJoyConDS4Report(leftSnapshot.buffer, rightSnapshot.buffer, dualPlayerPtr->gyroSource, MotionProfile::SwitchEmu);
                            dsuServer.UpdateController(dualPlayerPtr->dsuSlot, dsuReport);
                        }

                        auto ret = vigem_target_ds4_update_ex(vigem_client, dualPlayerPtr->ds4Controller, report);
                        const auto vigemComplete = SteadyClock::now();
                        if (!VIGEM_SUCCESS(ret))
                        {
                            std::wcerr << L"Failed to update DS4 report: 0x" << std::hex << ret << L"\n";
                        }

                        const auto newestReceivedAt = (leftSnapshot.receivedAt > rightSnapshot.receivedAt) ? leftSnapshot.receivedAt : rightSnapshot.receivedAt;
                        const double bleDeltaMs = (leftSnapshot.sequence >= rightSnapshot.sequence) ? leftSnapshot.bleDeltaMs : rightSnapshot.bleDeltaMs;
                        g_latencyLogger.Record(
                            g_runtimeOptions.updatePolicy,
                            ControllerTypeName(DualJoyCon),
                            ++dualSharedState->eventIndex,
                            bleDeltaMs,
                            MillisecondsBetween(leftSnapshot.receivedAt, vigemComplete),
                            MillisecondsBetween(rightSnapshot.receivedAt, vigemComplete),
                            MicrosecondsBetween(decodeStart, vigemComplete),
                            MicrosecondsBetween(newestReceivedAt, vigemComplete));
                    }
                });

            dualPlayers.push_back(std::move(dualPlayer));

            std::wcout << L"Dual Joy-Cons connected and configured. Press Enter to continue...\n";
            std::wstring dummy;
            std::getline(std::wcin, dummy);
        }
        else if (config.controllerType == ProController) {
            // Load or create GL/GR layouts configuration
            if (!LoadProControllerConfig(g_proControllerConfig)) {
                std::cout << "\nNo existing configuration found. Creating default layout...\n";
                CreateDefaultConfig();
                std::cout << "Default layout created: Layout 1 (GL: NONE, GR: NONE)\n";
                std::cout << "Press ZL+ZR+GL+GR during gameplay to open layout management.\n\n";
            } else {
                std::cout << "\nLoaded GL/GR layout configuration:\n";
                std::cout << "Active Layout: " << g_proControllerConfig.layouts[g_proControllerConfig.activeLayoutIndex].name << "\n";
                std::cout << "  GL -> " << ButtonMappingToString(g_proControllerConfig.layouts[g_proControllerConfig.activeLayoutIndex].glMapping) << "\n";
                std::cout << "  GR -> " << ButtonMappingToString(g_proControllerConfig.layouts[g_proControllerConfig.activeLayoutIndex].grMapping) << "\n";
                std::cout << "Total layouts: " << g_proControllerConfig.layouts.size() << "\n";
                std::cout << "Press ZL+ZR+GL+GR during gameplay to open layout management.\n";
                std::cout << "Press C button to cycle through layouts.\n\n";
            }

            std::wcout << L"Please sync your Pro Controller now.\n";

            ConnectedJoyCon proController = WaitForJoyCon(L"Waiting for Pro Controller...");

            // Request minimum BLE connection interval for lowest latency
            try {
                auto connectionParams = BluetoothLEPreferredConnectionParameters::ThroughputOptimized();
                auto paramResult = proController.device.RequestPreferredConnectionParameters(connectionParams);
                (void)paramResult;

                std::wcout << L"Requested ThroughputOptimized connection parameters for lower latency (Windows may ignore this request).\n";
            }
            catch (...) {
                std::wcout << L"Warning: Could not request preferred connection parameters. Continuing with default settings.\n";
            }

            PVIGEM_TARGET ds4_controller = vigem_target_ds4_alloc();
            auto ret = vigem_target_add(vigem_client, ds4_controller);
            if (!VIGEM_SUCCESS(ret))
            {
                std::wcerr << L"Failed to add DS4 controller target: 0x" << std::hex << ret << L"\n";
                exit(1);
            }

            auto proLatency = std::make_shared<LatencyTracker>();
            proController.inputChar.ValueChanged([ds4_controller, motionProfile = config.motionProfile, gyroMode = config.gyroMode, dsuSlot = static_cast<uint8_t>(i), &dsuServer, proLatency](GattCharacteristic const&, GattValueChangedEventArgs const& args) mutable
                {
                    const auto notificationReceived = SteadyClock::now();
                    auto reader = DataReader::FromBuffer(args.CharacteristicValue());
                    std::vector<uint8_t> buffer(reader.UnconsumedBufferLength());
                    reader.ReadBytes(buffer);
                    const double bleDeltaMs = MillisecondsBetween(proLatency->lastBleTime, notificationReceived);
                    proLatency->lastBleTime = notificationReceived;

                    if (!ShouldEmitForPolicy(g_runtimeOptions.updatePolicy, proLatency->lastEmitTime, SteadyClock::now())) {
                        return;
                    }

                    const auto decodeStart = SteadyClock::now();
                    const MotionProfile ds4Profile = (gyroMode == GyroMode::DsuUdp) ? MotionProfile::Raw : motionProfile;
                    DS4_REPORT_EX report = GenerateProControllerReport(buffer, ds4Profile);

                    // Apply GL/GR button mappings
                    ApplyGLGRMappings(report, buffer);

                    // Handle special buttons (Screenshot -> F12, C button)
                    HandleSpecialProButtons(buffer);

                    if (gyroMode == GyroMode::DsuUdp && dsuServer.IsRunning()) {
                        DS4_REPORT_EX dsuReport = GenerateProControllerReport(buffer, MotionProfile::SwitchEmu);
                        ApplyGLGRMappings(dsuReport, buffer);
                        dsuServer.UpdateController(dsuSlot, dsuReport);
                    }

                    auto ret = vigem_target_ds4_update_ex(vigem_client, ds4_controller, report);
                    const auto vigemComplete = SteadyClock::now();
                    if (!VIGEM_SUCCESS(ret)) {
                        std::wcerr << L"Failed to update DS4 EX report: 0x" << std::hex << ret << L"\n";
                    }

                    g_latencyLogger.Record(
                        g_runtimeOptions.updatePolicy,
                        ControllerTypeName(ProController),
                        ++proLatency->eventIndex,
                        bleDeltaMs,
                        0.0,
                        -1.0,
                        MicrosecondsBetween(decodeStart, vigemComplete),
                        MicrosecondsBetween(notificationReceived, vigemComplete));
                });

            auto status = proController.inputChar.WriteClientCharacteristicConfigurationDescriptorAsync(
                GattClientCharacteristicConfigurationDescriptorValue::Notify).get();

            if (proController.writeChar) {
                SendCustomCommands(proController.writeChar);
                std::this_thread::sleep_for(std::chrono::milliseconds(200));
                SetPlayerLEDs(proController.writeChar, 0x01);
                EmitSound(proController.writeChar);
            }

            if (status == GattCommunicationStatus::Success)
                std::wcout << L"Pro Controller notifications enabled.\n";
            else
                std::wcout << L"Failed to enable Pro Controller notifications.\n";

            if (config.gyroMode == GyroMode::DsuUdp && dsuServer.IsRunning()) {
                dsuServer.SetControllerConnected(static_cast<uint8_t>(i));
            }

            std::wcout << L"Press Enter to continue...\n";
            std::wstring dummy;
            std::getline(std::wcin, dummy);

            proPlayers.push_back({ proController, ds4_controller });
        }
        else if (config.controllerType == NSOGCController) {
            std::wcout << L"Please sync your NSO GameCube Controller now.\n";

            ConnectedJoyCon gcController = WaitForJoyCon(L"Waiting for NSO GC Controller...");

            // Request minimum BLE connection interval for lowest latency
            try {
                auto connectionParams = BluetoothLEPreferredConnectionParameters::ThroughputOptimized();
                gcController.device.RequestPreferredConnectionParameters(connectionParams);
                std::wcout << L"Requested ThroughputOptimized connection parameters for lower latency (Windows may ignore this request).\n";
            }
            catch (...) {
                std::wcout << L"Warning: Could not request preferred connection parameters.\n";
            }

            PVIGEM_TARGET ds4_controller = vigem_target_ds4_alloc();
            auto ret = vigem_target_add(vigem_client, ds4_controller);
            if (!VIGEM_SUCCESS(ret)) {
                std::wcerr << L"Failed to add DS4 controller target: 0x" << std::hex << ret << L"\n";
                exit(1);
            }

            gcController.inputChar.ValueChanged([ds4_controller](GattCharacteristic const&, GattValueChangedEventArgs const& args) mutable {
                auto reader = DataReader::FromBuffer(args.CharacteristicValue());
                std::vector<uint8_t> buffer(reader.UnconsumedBufferLength());
                reader.ReadBytes(buffer);

                DS4_REPORT_EX report = GenerateNSOGCReport(buffer);

                auto ret = vigem_target_ds4_update_ex(vigem_client, ds4_controller, report);
                if (!VIGEM_SUCCESS(ret)) {
                    std::wcerr << L"Failed to update DS4 EX report: 0x" << std::hex << ret << L"\n";
                }
                });

            auto status = gcController.inputChar.WriteClientCharacteristicConfigurationDescriptorAsync(
                GattClientCharacteristicConfigurationDescriptorValue::Notify).get();

            if (gcController.writeChar) {
                SendCustomCommands(gcController.writeChar); // Optional, only if NSO GC expects init commands
                std::this_thread::sleep_for(std::chrono::milliseconds(200));
                SetPlayerLEDs(gcController.writeChar, 0x01);
                EmitSound(gcController.writeChar);
            }

            if (status == GattCommunicationStatus::Success)
                std::wcout << L"NSO GC Controller notifications enabled.\n";
            else
                std::wcout << L"Failed to enable NSO GC Controller notifications.\n";

            std::wcout << L"Press Enter to continue...\n";
            std::wstring dummy;
            std::getline(std::wcin, dummy);

            proPlayers.push_back({ gcController, ds4_controller }); // reuse ProControllerPlayer struct
}
    }

    std::wcout << L"All players connected.\n";
    std::wcout << L"- Press Enter to exit\n";
    std::wcout << L"- Press ZL+ZR+GL+GR on Pro Controller to open layout management\n";
    std::wcout << L"- Press C button on Pro Controller to cycle layouts\n\n";

    // Main loop: monitor for management window trigger and exit command
    while (true) {
        // Check if management window should be opened
        if (g_openManagementWindow.load()) {
            g_openManagementWindow.store(false);
            std::cout << "\n[ZL+ZR+GL+GR combo detected! Opening layout management...]\n";
            ShowLayoutManagementWindow();
            std::cout << "\nResuming gameplay...\n";
            std::cout << "Active layout: " << g_proControllerConfig.layouts[g_proControllerConfig.activeLayoutIndex].name << "\n";
            std::cout << "Press Enter to exit, or ZL+ZR+GL+GR to manage layouts again.\n\n";
        }

        // Check for user input (non-blocking check)
        // Use a short sleep and check if Enter was pressed
        std::this_thread::sleep_for(std::chrono::milliseconds(100));

        // Check if input is available (simple approach for Windows)
        if (_kbhit()) {
            std::wstring dummy;
            std::getline(std::wcin, dummy);
            break;
        }
    }

    // Clean up dual player threads & free controllers
    for (auto& dp : dualPlayers)
    {
        dp->running.store(false);
        if (dp->sharedState) {
            dp->sharedState->cv.notify_all();
        }
        if (dp->updateThread.joinable())
            dp->updateThread.join();

        vigem_target_remove(vigem_client, dp->ds4Controller);
        vigem_target_free(dp->ds4Controller);
    }

    // Free single players controllers
    for (auto& sp : singlePlayers)
    {
        vigem_target_remove(vigem_client, sp.ds4Controller);
        vigem_target_free(sp.ds4Controller);
    }

    // Free Pro Controllers
    for (auto& pp : proPlayers)
    {
        vigem_target_remove(vigem_client, pp.ds4Controller);
        vigem_target_free(pp.ds4Controller);
    }

    if (vigem_client)
    {
        vigem_disconnect(vigem_client);
        vigem_free(vigem_client);
        vigem_client = nullptr;
    }

    dsuServer.Stop();

    return 0;
}
