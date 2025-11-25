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
#include <conio.h>  // For _kbhit()
#include "JoyConDecoder.h"
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
                            g_proControllerConfig.activeLayoutIndex = g_proControllerConfig.layouts.size() - 1;
                        }
                        SaveProControllerConfig(g_proControllerConfig);
                        std::cout << "Layout deleted!\n";
                    }
                } else {
                    std::cout << "Cannot delete the last layout!\n";
                }
            }
            else if (editChoice == "4") {
                g_proControllerConfig.activeLayoutIndex = index;
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

struct ConnectedJoyCon {
    BluetoothLEDevice device = nullptr;
    GattCharacteristic inputChar = nullptr;
    GattCharacteristic writeChar = nullptr;
};

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

    auto servicesResult = device.GetGattServicesAsync().get();
    if (servicesResult.Status() != GattCommunicationStatus::Success)
    {
        std::wcerr << L"Failed to get GATT services.\n";
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
};

// For single Joy-Con players, store controller + JoyCon info to keep alive
struct SingleJoyConPlayer {
    ConnectedJoyCon joycon;
    PVIGEM_TARGET ds4Controller;
    JoyConSide side;
    JoyConOrientation orientation;
};

// For dual Joy-Con players, store both JoyCons, controller, thread, and running flag
struct DualJoyConPlayer {
    ConnectedJoyCon leftJoyCon;
    ConnectedJoyCon rightJoyCon;
    GyroSource gyroSource;
    PVIGEM_TARGET ds4Controller;
    std::atomic<bool> running;
    std::thread updateThread;
};

// For Pro Controller players
struct ProControllerPlayer {
    ConnectedJoyCon controller;
    PVIGEM_TARGET ds4Controller;
};

// Declare the Pro Controller report generator (implement in JoyConDecoder.cpp)
DS4_REPORT_EX GenerateProControllerReport(const std::vector<uint8_t>& buffer);

int main()
{
    init_apartment();

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

        playerConfigs.push_back(config);
    }

    InitializeViGEm();

    // Store all players to keep them alive
    std::vector<SingleJoyConPlayer> singlePlayers;
    std::vector<std::unique_ptr<DualJoyConPlayer>> dualPlayers;
    std::vector<ProControllerPlayer> proPlayers;

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
                std::wcout << L"Requested ThroughputOptimized connection parameters for lower latency.\n";
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

            player.joycon.inputChar.ValueChanged([joyconSide = player.side, joyconOrientation = player.orientation, &player](GattCharacteristic const&, GattValueChangedEventArgs const& args)
                {
                    auto reader = DataReader::FromBuffer(args.CharacteristicValue());
                    std::vector<uint8_t> buffer(reader.UnconsumedBufferLength());
                    reader.ReadBytes(buffer);

                    DS4_REPORT_EX report = GenerateDS4Report(buffer, joyconSide, joyconOrientation);

                    auto ret = vigem_target_ds4_update_ex(vigem_client, player.ds4Controller, report);
                    if (!VIGEM_SUCCESS(ret)) {
                        std::wcerr << L"Failed to update DS4 EX report: 0x" << std::hex << ret << L"\n";
                    }
                });

            auto status = player.joycon.inputChar.WriteClientCharacteristicConfigurationDescriptorAsync(
                GattClientCharacteristicConfigurationDescriptorValue::Notify).get();

            if (player.joycon.writeChar)
                SendCustomCommands(player.joycon.writeChar);

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
            if (rightJoyCon.writeChar)
                SendCustomCommands(rightJoyCon.writeChar);

            // Request minimum BLE connection interval for right Joy-Con
            try {
                auto connectionParams = BluetoothLEPreferredConnectionParameters::ThroughputOptimized();
                rightJoyCon.device.RequestPreferredConnectionParameters(connectionParams);
                std::wcout << L"Requested ThroughputOptimized connection parameters for RIGHT Joy-Con.\n";
            }
            catch (...) {
                std::wcout << L"Warning: Could not request preferred connection parameters for RIGHT Joy-Con.\n";
            }

            std::wcout << L"Please sync your LEFT Joy-Con now.\n";
            ConnectedJoyCon leftJoyCon = WaitForJoyCon(L"Waiting for LEFT Joy-Con...");
            if (leftJoyCon.writeChar)
                SendCustomCommands(leftJoyCon.writeChar);

            // Request minimum BLE connection interval for left Joy-Con
            try {
                auto connectionParams = BluetoothLEPreferredConnectionParameters::ThroughputOptimized();
                leftJoyCon.device.RequestPreferredConnectionParameters(connectionParams);
                std::wcout << L"Requested ThroughputOptimized connection parameters for LEFT Joy-Con.\n";
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
            dualPlayer->ds4Controller = ds4Controller;
            dualPlayer->running.store(true);

            std::atomic<std::shared_ptr<std::vector<uint8_t>>> leftBufferAtomic{ std::make_shared<std::vector<uint8_t>>() };
            std::atomic<std::shared_ptr<std::vector<uint8_t>>> rightBufferAtomic{ std::make_shared<std::vector<uint8_t>>() };

            dualPlayer->leftJoyCon.inputChar.ValueChanged([&leftBufferAtomic](GattCharacteristic const&, GattValueChangedEventArgs const& args)
                {
                    auto reader = DataReader::FromBuffer(args.CharacteristicValue());
                    auto buf = std::make_shared<std::vector<uint8_t>>(reader.UnconsumedBufferLength());
                    reader.ReadBytes(*buf);
                    leftBufferAtomic.store(buf, std::memory_order_release);
                });

            auto statusLeft = dualPlayer->leftJoyCon.inputChar.WriteClientCharacteristicConfigurationDescriptorAsync(
                GattClientCharacteristicConfigurationDescriptorValue::Notify).get();

            if (statusLeft == GattCommunicationStatus::Success)
                std::wcout << L"LEFT Joy-Con notifications enabled.\n";
            else
                std::wcout << L"Failed to enable LEFT Joy-Con notifications.\n";

            dualPlayer->rightJoyCon.inputChar.ValueChanged([&rightBufferAtomic](GattCharacteristic const&, GattValueChangedEventArgs const& args)
                {
                    auto reader = DataReader::FromBuffer(args.CharacteristicValue());
                    auto buf = std::make_shared<std::vector<uint8_t>>(reader.UnconsumedBufferLength());
                    reader.ReadBytes(*buf);
                    rightBufferAtomic.store(buf, std::memory_order_release);
                });

            auto statusRight = dualPlayer->rightJoyCon.inputChar.WriteClientCharacteristicConfigurationDescriptorAsync(
                GattClientCharacteristicConfigurationDescriptorValue::Notify).get();

            if (statusRight == GattCommunicationStatus::Success)
                std::wcout << L"RIGHT Joy-Con notifications enabled.\n";
            else
                std::wcout << L"Failed to enable RIGHT Joy-Con notifications.\n";

            dualPlayer->updateThread = std::thread([dualPlayerPtr = dualPlayer.get(), &leftBufferAtomic, &rightBufferAtomic]()
                {
                    while (dualPlayerPtr->running.load(std::memory_order_acquire))
                    {
                        auto leftBuf = leftBufferAtomic.load(std::memory_order_acquire);
                        auto rightBuf = rightBufferAtomic.load(std::memory_order_acquire);

                        if (leftBuf->empty() || rightBuf->empty())
                        {
                            std::this_thread::sleep_for(std::chrono::milliseconds(5));
                            continue;
                        }

                        DS4_REPORT_EX report = GenerateDualJoyConDS4Report(*leftBuf, *rightBuf, dualPlayerPtr->gyroSource);

                        auto ret = vigem_target_ds4_update_ex(vigem_client, dualPlayerPtr->ds4Controller, report);
                        if (!VIGEM_SUCCESS(ret))
                        {
                            std::wcerr << L"Failed to update DS4 report: 0x" << std::hex << ret << L"\n";
                        }

                        std::this_thread::sleep_for(std::chrono::milliseconds(16)); // ~60Hz
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

                std::wcout << L"Requested ThroughputOptimized connection parameters for lower latency.\n";
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

            proController.inputChar.ValueChanged([ds4_controller](GattCharacteristic const&, GattValueChangedEventArgs const& args) mutable
                {
                    auto reader = DataReader::FromBuffer(args.CharacteristicValue());
                    std::vector<uint8_t> buffer(reader.UnconsumedBufferLength());
                    reader.ReadBytes(buffer);


                    DS4_REPORT_EX report = GenerateProControllerReport(buffer);

                    // Apply GL/GR button mappings
                    ApplyGLGRMappings(report, buffer);

                    // Handle special buttons (Screenshot -> F12, C button)
                    HandleSpecialProButtons(buffer);

                    auto ret = vigem_target_ds4_update_ex(vigem_client, ds4_controller, report);
                    if (!VIGEM_SUCCESS(ret)) {
                        std::wcerr << L"Failed to update DS4 EX report: 0x" << std::hex << ret << L"\n";
                    }
                });

            auto status = proController.inputChar.WriteClientCharacteristicConfigurationDescriptorAsync(
                GattClientCharacteristicConfigurationDescriptorValue::Notify).get();

            if (proController.writeChar)
                SendCustomCommands(proController.writeChar);

            if (status == GattCommunicationStatus::Success)
                std::wcout << L"Pro Controller notifications enabled.\n";
            else
                std::wcout << L"Failed to enable Pro Controller notifications.\n";

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
                std::wcout << L"Requested ThroughputOptimized connection parameters for lower latency.\n";
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

            if (gcController.writeChar)
                SendCustomCommands(gcController.writeChar); // Optional, only if NSO GC expects init commands

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

    return 0;
}
