#define NOMINMAX

#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/Windows.Storage.Streams.h>
#include <winrt/Windows.Devices.Bluetooth.h>
#include <winrt/Windows.Devices.Bluetooth.Advertisement.h>
#include <winrt/Windows.Devices.Bluetooth.GenericAttributeProfile.h>

#include <d3d11.h>
#include <dxgi.h>
#include <tchar.h>

#include "imgui.h"
#include "imgui_impl_win32.h"
#include "imgui_impl_dx11.h"

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
#include <functional>

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

extern IMGUI_IMPL_API LRESULT ImGui_ImplWin32_WndProcHandler(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);

constexpr uint16_t JOYCON_MANUFACTURER_ID = 1363;
const std::vector<uint8_t> JOYCON_MANUFACTURER_PREFIX = { 0x01, 0x00, 0x03, 0x7E };
const wchar_t* INPUT_REPORT_UUID  = L"ab7de9be-89fe-49ad-828f-118f09df7fd2";
const wchar_t* WRITE_COMMAND_UUID = L"649d4ac9-8eb7-4e6c-af44-1ea54fe5f005";


const wchar_t* CMD_RESPONSE_UUID_BASIC = L"c765a961-d9d8-4d36-a20a-5315b111836a";
const wchar_t* CMD_RESPONSE_UUID_L     = L"63a3810f-aec7-474b-9010-3d52403cb996";
const wchar_t* CMD_RESPONSE_UUID_R     = L"640ca58e-0e88-410c-a7f3-426faf2b690b";

const wchar_t* RUMBLE_CHAR_UUID_L = L"ce49a830-dced-48ae-931e-c8cf88aadbea";
const wchar_t* RUMBLE_CHAR_UUID_R = L"65a724b3-f1e7-4a61-8078-a342376b27ff";

const wchar_t* VIBRATION_CHAR_UUID_L = L"289326cb-a471-485d-a8f4-240c14f18241";
const wchar_t* VIBRATION_CHAR_UUID_R = L"fa19b0fb-cd1f-46a7-84a1-bbb09e00c149";

const std::string CONFIG_FILE = "joycon2cpp_config.json";

enum class UpdatePolicy { LowLatency, Balanced120Hz, Legacy60Hz };
enum ControllerType { SingleJoyCon = 1, DualJoyCon = 2, ProController = 3, NSOGCController = 4 };

enum class ButtonMapping {
    NONE, L3, R3, L1, R1, L2, R2,
    CROSS, CIRCLE, SQUARE, TRIANGLE,
    SHARE, OPTIONS,
    DPAD_UP, DPAD_DOWN, DPAD_LEFT, DPAD_RIGHT
};

enum class AppScreen { Setup, Connecting, Running };

struct GLGRLayout {
    char name[64] = "Layout 1";
    ButtonMapping glMapping = ButtonMapping::NONE;
    ButtonMapping grMapping = ButtonMapping::NONE;
};

struct ProControllerConfig {
    std::vector<GLGRLayout> layouts;
    int activeLayoutIndex = 0;
};

struct RuntimeOptions {
    bool latencyMetrics = false;
    UpdatePolicy updatePolicy = UpdatePolicy::LowLatency;
    char latencyCsvPath[256] = "latency_benchmark.csv";
};

struct PlayerConfig {
    ControllerType controllerType = SingleJoyCon;
    JoyConSide     joyconSide        = JoyConSide::Left;
    JoyConOrientation joyconOrientation = JoyConOrientation::Upright;
    GyroSource     gyroSource        = GyroSource::Both;
    GyroMode       gyroMode          = GyroMode::Raw;
};

struct ConnectedJoyCon {
    BluetoothLEDevice      device    = nullptr;
    GattCharacteristic     inputChar = nullptr;
    GattCharacteristic     writeChar = nullptr;
    GattCharacteristic     rumbleChar = nullptr;
    GattCharacteristic     vibrationChar = nullptr;
    GattCharacteristic     vibrationCharLeft = nullptr;
    GattCharacteristic     vibrationCharRight = nullptr;
    GattCharacteristic     cmdRespBasicChar = nullptr;
    GattCharacteristic     cmdRespExtChar = nullptr;
};

using SteadyClock = std::chrono::steady_clock;
using TimePoint   = SteadyClock::time_point;

struct LatencyTracker {
    TimePoint lastBleTime{};
    TimePoint lastEmitTime{};
    uint64_t  eventIndex = 0;
};

struct TimedInputBuffer {
    std::vector<uint8_t> buffer;
    TimePoint  receivedAt{};
    double     bleDeltaMs = -1.0;
    uint64_t   sequence   = 0;
};

struct DualJoyConSharedState {
    std::mutex              mutex;
    std::condition_variable cv;
    TimedInputBuffer        left, right;
    TimePoint               lastLeftBleTime{}, lastRightBleTime{}, lastEmitTime{};
    uint64_t                sequence = 0, eventIndex = 0;
};

struct SingleJoyConPlayer {
    ConnectedJoyCon joycon;
    PVIGEM_TARGET   ds4Controller = nullptr;
    JoyConSide      side;
    JoyConOrientation orientation;
    int             mouseMode      = 0;
    bool            wasChatPressed = false;
    int16_t         lastOpticalX   = 0, lastOpticalY = 0;
    bool            firstOpticalRead = true;
    float           scrollAccumulator = 0.f;
    bool            mb4Pressed = false, mb5Pressed = false;
    bool            leftBtnPressed = false, rightBtnPressed = false, middleBtnPressed = false;
    LatencyTracker  latency;
};

struct DualJoyConPlayer {
    ConnectedJoyCon leftJoyCon, rightJoyCon;
    GyroSource      gyroSource;
    GyroMode        gyroMode;
    uint8_t         dsuSlot = 0;
    PVIGEM_TARGET   ds4Controller = nullptr;
    std::atomic<bool> running{false};
    std::thread     updateThread;
    std::shared_ptr<DualJoyConSharedState> sharedState;
};

struct ProControllerPlayer {
    ConnectedJoyCon controller;
    PVIGEM_TARGET   ds4Controller = nullptr;
    LatencyTracker  latency;
};

struct ConnectionTask {
    std::string        label;
    bool               done    = false;
    bool               success = false;
    std::string        statusMsg;
    ConnectedJoyCon    result;
    std::thread        worker;
};


struct SingleRumbleCtx {
    GattCharacteristic  vibrationChar;
    PVIGEM_TARGET       target    = nullptr;
    GattCharacteristic  secondaryVibrationChar = nullptr;
    std::atomic<bool>   running{true};
    std::thread         thread;
};
struct DualRumbleCtx {
    GattCharacteristic  leftVibration;
    GattCharacteristic  rightVibration;
    PVIGEM_TARGET       target    = nullptr;
    std::atomic<bool>   running{true};
    std::thread         thread;
};

static AppScreen              g_screen        = AppScreen::Setup;
static RuntimeOptions         g_opts;
static ProControllerConfig    g_proConfig;
static PVIGEM_CLIENT          g_vigem         = nullptr;
static DsuServer              g_dsuServer;

static std::vector<PlayerConfig>                    g_playerConfigs;
static std::vector<SingleJoyConPlayer>              g_singlePlayers;
static std::vector<std::unique_ptr<DualJoyConPlayer>> g_dualPlayers;
static std::vector<ProControllerPlayer>             g_proPlayers;

static std::vector<SingleRumbleCtx*> g_singleRumbleCtxs;
static std::vector<DualRumbleCtx*>   g_dualRumbleCtxs;
static std::vector<SingleRumbleCtx*> g_proRumbleCtxs;

static std::vector<ConnectionTask>  g_connectionTasks;
static int                          g_connectionTaskIndex = 0;
static bool                         g_connectionDone      = false;
static std::string                  g_connectionError;

static bool g_showLayoutManager = false;
static bool g_screenshotButtonPressed = false;
static bool g_cButtonPressed          = false;
static bool g_comboPressed            = false;
static std::atomic<bool> g_openLayoutManager{false};
static std::atomic<bool> g_shuttingDown{false};

static std::mutex         g_logMutex;
static std::vector<std::string> g_logLines;
static void AppLog(const std::string& s) {
    std::lock_guard<std::mutex> lk(g_logMutex);
    g_logLines.push_back(s);
    if (g_logLines.size() > 200) g_logLines.erase(g_logLines.begin());
}

class LatencyCsvLogger {
public:
    bool Start(const std::string& path) {
        std::lock_guard<std::mutex> lk(m_mutex);
        m_file.open(path, std::ios::out | std::ios::trunc);
        if (!m_file.is_open()) return false;
        m_file << "mode,controller_type,event_index,ble_delta_ms,buffer_age_left_ms,buffer_age_right_ms,decode_to_vigem_us,total_pipeline_us\n";
        m_enabled = true;
        return true;
    }
    bool Enabled() const { return m_enabled.load(std::memory_order_acquire); }
    void Record(UpdatePolicy pol, const char* ct, uint64_t idx,
                double bleDelta, double ageL, double ageR, double decUs, double totUs) {
        if (!Enabled()) return;
        std::lock_guard<std::mutex> lk(m_mutex);
        if (!m_file.is_open()) return;
        m_file << PolicyName(pol) << ',' << ct << ',' << idx << ','
               << std::fixed << std::setprecision(3)
               << bleDelta << ',' << ageL << ',' << ageR << ','
               << std::setprecision(1) << decUs << ',' << totUs << '\n';
    }
    static const char* PolicyName(UpdatePolicy p) {
        switch(p) {
            case UpdatePolicy::LowLatency:    return "LowLatency";
            case UpdatePolicy::Balanced120Hz: return "Balanced120Hz";
            case UpdatePolicy::Legacy60Hz:    return "Legacy60Hz";
            default: return "Unknown";
        }
    }
private:
    std::atomic<bool> m_enabled{false};
    std::mutex        m_mutex;
    std::ofstream     m_file;
};
static LatencyCsvLogger g_latencyLogger;

static double MsBetween(TimePoint a, TimePoint b) {
    if (a == TimePoint{}) return -1.0;
    return std::chrono::duration<double, std::milli>(b - a).count();
}
static double UsBetween(TimePoint a, TimePoint b) {
    return std::chrono::duration<double, std::micro>(b - a).count();
}
static std::chrono::microseconds PolicyInterval(UpdatePolicy p) {
    switch(p) {
        case UpdatePolicy::Balanced120Hz: return std::chrono::microseconds(8333);
        case UpdatePolicy::Legacy60Hz:    return std::chrono::microseconds(16667);
        default:                          return std::chrono::microseconds(0);
    }
}
static bool ShouldEmit(UpdatePolicy p, TimePoint& last, TimePoint now) {
    auto iv = PolicyInterval(p);
    if (iv.count() == 0 || last == TimePoint{} || now - last >= iv) { last = now; return true; }
    return false;
}
static const char* CtrlTypeName(int t) {
    switch(t) {
        case 1: return "SingleJoyCon";
        case 2: return "DualJoyCon";
        case 3: return "ProController";
        case 4: return "NSOGCController";
        default: return "Unknown";
    }
}

static const char* BtnMapNames[] = {
    "None","L3","R3","L1","R1","L2","R2",
    "Cross","Circle","Square","Triangle",
    "Share","Options",
    "DPad Up","DPad Down","DPad Left","DPad Right"
};
static const char* BtnMapStr(ButtonMapping m) { return BtnMapNames[(int)m]; }
static ButtonMapping BtnMapFromStr(const std::string& s) {
    for (int i = 0; i < 17; i++)
        if (s == BtnMapNames[i]) return (ButtonMapping)i;
    return ButtonMapping::NONE;
}

static void SaveProConfig() {
    std::ofstream f(CONFIG_FILE);
    if (!f.is_open()) return;
    f << "{\n  \"activeLayoutIndex\": " << g_proConfig.activeLayoutIndex << ",\n  \"layouts\": [\n";
    for (size_t i = 0; i < g_proConfig.layouts.size(); ++i) {
        auto& l = g_proConfig.layouts[i];
        f << "    {\"name\":\"" << l.name << "\","
          << "\"glMapping\":\"" << BtnMapStr(l.glMapping) << "\","
          << "\"grMapping\":\"" << BtnMapStr(l.grMapping) << "\"}";
        if (i+1 < g_proConfig.layouts.size()) f << ",";
        f << "\n";
    }
    f << "  ]\n}\n";
}
static void LoadProConfig() {
    std::ifstream f(CONFIG_FILE);
    if (!f.is_open()) {
        GLGRLayout def; def.glMapping = ButtonMapping::NONE; def.grMapping = ButtonMapping::NONE;
        g_proConfig.layouts.push_back(def);
        g_proConfig.activeLayoutIndex = 0;
        SaveProConfig();
        return;
    }
    std::string json((std::istreambuf_iterator<char>(f)), std::istreambuf_iterator<char>());
    g_proConfig.layouts.clear(); g_proConfig.activeLayoutIndex = 0;
    auto readField = [&](const std::string& key) -> std::string {
        auto p = json.find("\"" + key + "\"");
        if (p == std::string::npos) return "";
        auto c = json.find(':', p); auto e = json.find_first_of(",}\n", c);
        auto v = json.substr(c+1, e-c-1);
        v.erase(0, v.find_first_not_of(" \t\n\r\""));
        v.erase(v.find_last_not_of(" \t\n\r\"")+1);
        return v;
    };
    try { g_proConfig.activeLayoutIndex = std::stoi(readField("activeLayoutIndex")); } catch(...) {}
    size_t pos = json.find('[');
    while (pos != std::string::npos) {
        auto ob = json.find('{', pos);
        if (ob == std::string::npos) break;
        auto oe = json.find('}', ob);
        if (oe == std::string::npos) break;
        auto obj = json.substr(ob, oe-ob+1);
        GLGRLayout lay;
        auto getStr = [&](const std::string& key) -> std::string {
            auto kp = obj.find("\""+key+"\"");
            if (kp == std::string::npos) return "";
            auto q1 = obj.find('"', kp+key.size()+2);
            auto q2 = obj.find('"', q1+1);
            return obj.substr(q1+1, q2-q1-1);
        };
        auto nm = getStr("name");
        if (!nm.empty()) strncpy_s(lay.name, nm.c_str(), 63);
        lay.glMapping = BtnMapFromStr(getStr("glMapping"));
        lay.grMapping = BtnMapFromStr(getStr("grMapping"));
        g_proConfig.layouts.push_back(lay);
        pos = oe+1;
    }
    if (g_proConfig.layouts.empty()) {
        GLGRLayout def; g_proConfig.layouts.push_back(def); g_proConfig.activeLayoutIndex = 0;
    }
    if (g_proConfig.activeLayoutIndex < 0 || g_proConfig.activeLayoutIndex >= (int)g_proConfig.layouts.size())
        g_proConfig.activeLayoutIndex = 0;
}


static std::string HexBytes(const std::vector<uint8_t>& bytes) {
    std::string hex;
    hex.reserve(bytes.size() * 3);
    for (auto b : bytes) {
        char buf[4];
        sprintf_s(buf, "%02X ", b);
        hex += buf;
    }
    return hex;
}

static void AttachCommandResponseLogger(
    GattCharacteristic const& ch,
    const std::string& name,
    GattClientCharacteristicConfigurationDescriptorValue cccValue =
        GattClientCharacteristicConfigurationDescriptorValue::Notify)
{
    if (!ch) {
        return;
    }

    ch.ValueChanged([name](GattCharacteristic const&, GattValueChangedEventArgs const& args) {
        if (g_shuttingDown.load()) return;
        auto rdr = DataReader::FromBuffer(args.CharacteristicValue());
        std::vector<uint8_t> buf(rdr.UnconsumedBufferLength());
        rdr.ReadBytes(buf);


        if (buf.size() >= 8) {
            char tmp[160];
            sprintf_s(tmp,"[CMDRESP] parsed cmd=0x%02X sub=0x%02X ack0=0x%02X ack1=0x%02X transport=0x%02X",buf[0], buf[3], buf[4], buf[5], buf[2]);
        }
    });

    try {
        auto status = ch.WriteClientCharacteristicConfigurationDescriptorAsync(cccValue).get();
    } catch (...) { }
}

static void SendGenericCommand(GattCharacteristic const& ch, uint8_t cmdId, uint8_t subId, const std::vector<uint8_t>& data) {
    if (g_shuttingDown.load()) return;
    if (!ch) {
        return;
    }

    std::vector<uint8_t> pkt;
    pkt.reserve(8 + data.size());
    pkt.push_back(cmdId);
    pkt.push_back(0x91);
    pkt.push_back(0x01); 
    pkt.push_back(subId);
    pkt.push_back(0x00);
    pkt.push_back((uint8_t)data.size());
    pkt.push_back(0x00);
    pkt.push_back(0x00);
    pkt.insert(pkt.end(), data.begin(), data.end());

    DataWriter w;
    w.WriteBytes(pkt);
    try {
        auto status = ch.WriteValueAsync(
            w.DetachBuffer(),
            GattWriteOption::WriteWithoutResponse
        ).get();
    } catch (...) {
        return;
    }
    std::this_thread::sleep_for(std::chrono::milliseconds(50));
}
static void SetPlayerLEDs(GattCharacteristic const& ch, uint8_t pat) {
    std::vector<uint8_t> d(8, 0); d[0] = pat;
    SendGenericCommand(ch, 0x09, 0x07, d);
}

static void SendVibrationSample(GattCharacteristic const& ch, uint8_t sampleId);

static void EmitSound(GattCharacteristic const& ch) {
    SendVibrationSample(ch, 0x04);
}
static void SendCustomCommands(GattCharacteristic const& ch) {
    if (!ch || g_shuttingDown.load()) return;
    std::vector<std::vector<uint8_t>> cmds = {
        {0x0c,0x91,0x01,0x02,0x00,0x04,0x00,0x00,0x37,0x00,0x00,0x00},

        {0x0c,0x91,0x01,0x04,0x00,0x04,0x00,0x00,0x37,0x00,0x00,0x00},
    };

    for (auto& cmd : cmds) {
        DataWriter w;
        w.WriteBytes(cmd);

        try {
            auto status = ch.WriteValueAsync(
                w.DetachBuffer(),
                GattWriteOption::WriteWithoutResponse
            ).get();
        } catch (...) {
            return;
        }
        std::this_thread::sleep_for(std::chrono::milliseconds(100));
    }
}

static void Send0016Command(GattCharacteristic const& ch, const std::vector<uint8_t>& cmd)
{
    if (!ch || g_shuttingDown.load()) return;

    std::vector<uint8_t> pkt(17, 0x00);
    pkt.insert(pkt.end(), cmd.begin(), cmd.end());

    DataWriter w;
    w.WriteBytes(pkt);
    ch.WriteValueAsync(w.DetachBuffer(), GattWriteOption::WriteWithoutResponse);
}


static void SendJoyCon2OfficialInit(GattCharacteristic const& ch)
{
    if (!ch || g_shuttingDown.load()) return;

    std::vector<std::vector<uint8_t>> cmds = {
        {0x07,0x91,0x01,0x01,0x00,0x00,0x00,0x00},
        {0x02,0x91,0x01,0x04,0x00,0x08,0x00,0x00,0x40,0x7E,0x00,0x00,0x00,0x30,0x01,0x00},
        {0x10,0x91,0x01,0x01,0x00,0x00,0x00,0x00},
        {0x16,0x91,0x01,0x01,0x00,0x00,0x00,0x00},

        {0x0A,0x91,0x01,0x02,0x00,0x04,0x00,0x00,0x03,0x00,0x00,0x00},
        {0x09,0x91,0x01,0x07,0x00,0x08,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00},
        {0x0C,0x91,0x01,0x02,0x00,0x04,0x00,0x00,0x37,0x00,0x00,0x00},

        {0x02,0x91,0x01,0x04,0x00,0x08,0x00,0x00,0x40,0x7E,0x00,0x00,0x80,0x30,0x01,0x00},
        {0x02,0x91,0x01,0x04,0x00,0x08,0x00,0x00,0x40,0x7E,0x00,0x00,0x40,0xC0,0x1F,0x00},
        {0x02,0x91,0x01,0x04,0x00,0x08,0x00,0x00,0x10,0x7E,0x00,0x00,0x40,0x30,0x01,0x00},
        {0x02,0x91,0x01,0x04,0x00,0x08,0x00,0x00,0x18,0x7E,0x00,0x00,0x00,0x31,0x01,0x00},
        {0x11,0x91,0x01,0x03,0x00,0x00,0x00,0x00},
        {0x02,0x91,0x01,0x04,0x00,0x08,0x00,0x00,0x20,0x7E,0x00,0x00,0x60,0x30,0x01,0x00},


        {0x0A,0x91,0x01,0x08,0x00,0x14,0x00,0x00,
         0x01,0x59,0x09,0x00,0x00,0xFF,0xFF,0xFF,0xFF,0x35,0x00,0x46,
         0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00},

        {0x11,0x91,0x01,0x01,0x00,0x00,0x00,0x00},
        {0x0C,0x91,0x01,0x04,0x00,0x04,0x00,0x00,0x37,0x00,0x00,0x00},
    };

    for (auto& cmd : cmds) {
        Send0016Command(ch, cmd);
        std::this_thread::sleep_for(std::chrono::milliseconds(80));
    }
}

static void SendRumble(GattCharacteristic const& ch, float amp)
{
    if (!ch || g_shuttingDown.load()) return;

    static uint8_t c = 0;

    std::vector<uint8_t> pkt(42, 0x00);
    pkt[0] = 0x00;

    if (amp >= 0.01f) {
        pkt[1] = 0x50 | (c & 0x0F);

        const uint8_t strong[5] = {0x93, 0x35, 0x36, 0x1C, 0x0D};
        const uint8_t weak[5]   = {0x4B, 0x7D, 0x80, 0x5A, 0x02};

        const uint8_t* p = amp > 0.5f ? strong : weak;
        memcpy(pkt.data() + 2, p, 5);

        c = (c + 1) & 0x0F;
    }

    DataWriter w;
    w.WriteBytes(pkt);
    ch.WriteValueAsync(w.DetachBuffer(), GattWriteOption::WriteWithoutResponse);
}

static void SendVibrationSample(GattCharacteristic const& ch, uint8_t sampleId)
{
    std::vector<uint8_t> data(4, 0);
    data[0] = sampleId;
    SendGenericCommand(ch, 0x0A, 0x02, data);
}

static void StartSingleRumbleThread(SingleRumbleCtx* ctx) {
    ctx->thread = std::thread([ctx]() {
        DS4_OUTPUT_DATA last{};
        int silentFrames = 0;
        while (ctx->running.load() && !g_shuttingDown.load()) {
            DS4_OUTPUT_DATA out{};
            auto err = vigem_target_ds4_get_output(g_vigem, ctx->target, &out);
            if (VIGEM_SUCCESS(err)) {
                float amp = std::max(out.LargeMotor, out.SmallMotor) / 255.0f;
                if (amp >= 0.01f) {
                    SendRumble(ctx->vibrationChar, amp);
                    if (ctx->secondaryVibrationChar)
                        SendRumble(ctx->secondaryVibrationChar, amp);
                    silentFrames = 0;
                } else {
                    if (silentFrames == 0) {
                        SendRumble(ctx->vibrationChar, 0.0f);
                        if (ctx->secondaryVibrationChar)
                            SendRumble(ctx->secondaryVibrationChar, 0.0f);
                    }
                    silentFrames++;
                }
                last = out;
            }
            std::this_thread::sleep_for(std::chrono::milliseconds(4));
        }
    });
}

static void StartDualRumbleThread(DualRumbleCtx* ctx) {
    ctx->thread = std::thread([ctx]() {
        DS4_OUTPUT_DATA last{};
        int silentFrames = 0;
        while (ctx->running.load() && !g_shuttingDown.load()) {
            DS4_OUTPUT_DATA out{};
            auto err = vigem_target_ds4_get_output(g_vigem, ctx->target, &out);
            if (VIGEM_SUCCESS(err)) {
                float amp = std::max(out.LargeMotor, out.SmallMotor) / 255.0f;
                if (amp >= 0.01f) {
                    SendRumble(ctx->leftVibration, amp);
                    SendRumble(ctx->rightVibration, amp);
                    silentFrames = 0;
                } else {
                    if (silentFrames == 0) {
                        SendRumble(ctx->leftVibration, 0.0f);
                        SendRumble(ctx->rightVibration, 0.0f);
                    }
                    silentFrames++;
                }
                last = out;
            }
            std::this_thread::sleep_for(std::chrono::milliseconds(4));
        }
    });
}


static ConnectedJoyCon BleConnect(std::function<void(const std::string&)> statusCb, bool& outSuccess, JoyConSide side = JoyConSide::Left) {
    outSuccess = false;
    ConnectedJoyCon cj{};
    BluetoothLEDevice device = nullptr;
    bool found = false;
    BluetoothLEAdvertisementWatcher watcher;
    std::mutex mtx; std::condition_variable cv;

    watcher.Received([&](auto const&, auto const& args) {
        if (g_shuttingDown.load()) return;
        std::unique_lock<std::mutex> lk(mtx);
        if (found) return;
        auto mfg = args.Advertisement().ManufacturerData();
        for (uint32_t i = 0; i < mfg.Size(); i++) {
            auto sec = mfg.GetAt(i);
            if (sec.CompanyId() != JOYCON_MANUFACTURER_ID) continue;
            auto rdr = DataReader::FromBuffer(sec.Data());
            std::vector<uint8_t> d(rdr.UnconsumedBufferLength());
            rdr.ReadBytes(d);
            if (d.size() >= JOYCON_MANUFACTURER_PREFIX.size() &&
                std::equal(JOYCON_MANUFACTURER_PREFIX.begin(), JOYCON_MANUFACTURER_PREFIX.end(), d.begin())) {
                device = BluetoothLEDevice::FromBluetoothAddressAsync(args.BluetoothAddress()).get();
                if (!device) return;
                found = true; watcher.Stop(); cv.notify_one();
            }
        }
    });
    watcher.ScanningMode(BluetoothLEScanningMode::Active);
    watcher.Start();
    statusCb("Scanning... (hold sync button)");
    {
        std::unique_lock<std::mutex> lk(mtx);
        if (!cv.wait_for(lk, std::chrono::seconds(30), [&]{ return found || g_shuttingDown.load(); })) {
            watcher.Stop();
            statusCb(g_shuttingDown.load() ? "Cancelled" : "Timed out — device not found");
            return cj;
        }
        if (g_shuttingDown.load()) { watcher.Stop(); statusCb("Cancelled"); return cj; }
    }
    statusCb("Device found, fetching GATT services...");
    cj.device = device;
    GattDeviceServicesResult sr = nullptr;
    for (int att = 1; att <= 10; ++att) {
        sr = device.GetGattServicesAsync(BluetoothCacheMode::Uncached).get();
        if (sr.Status() == GattCommunicationStatus::Success) break;
        statusCb("Retrying GATT (" + std::to_string(att) + "/10)...");
        std::this_thread::sleep_for(std::chrono::milliseconds(500));
    }
    if (!sr || sr.Status() != GattCommunicationStatus::Success) {
        statusCb("Failed to get GATT services. Re-pair the device.");
        return cj;
    }
    for (auto svc : sr.Services()) {
        auto cr = svc.GetCharacteristicsAsync().get();
        if (cr.Status() != GattCommunicationStatus::Success) continue;
        for (auto ch : cr.Characteristics()) {
            if (ch.Uuid() == guid(INPUT_REPORT_UUID))  cj.inputChar  = ch;
            if (ch.Uuid() == guid(WRITE_COMMAND_UUID)) cj.writeChar  = ch;
            if (ch.Uuid() == guid(CMD_RESPONSE_UUID_BASIC)) cj.cmdRespBasicChar = ch;
            if (ch.Uuid() == guid(CMD_RESPONSE_UUID_L)) { if (side == JoyConSide::Left)  cj.cmdRespExtChar = ch; }
            if (ch.Uuid() == guid(CMD_RESPONSE_UUID_R)) { if (side == JoyConSide::Right) cj.cmdRespExtChar = ch; }
            if (ch.Uuid() == guid(RUMBLE_CHAR_UUID_L)) { if (side == JoyConSide::Left)  cj.rumbleChar = ch; }
            if (ch.Uuid() == guid(RUMBLE_CHAR_UUID_R)) { if (side == JoyConSide::Right) cj.rumbleChar = ch; }
            if (ch.Uuid() == guid(VIBRATION_CHAR_UUID_L)) {
                cj.vibrationCharLeft = ch;
                if (side == JoyConSide::Left) cj.vibrationChar = ch;
            }
            if (ch.Uuid() == guid(VIBRATION_CHAR_UUID_R)) {
                cj.vibrationCharRight = ch;
                if (side == JoyConSide::Right) cj.vibrationChar = ch;
            }
        }
    }
    if (!cj.inputChar || !cj.writeChar) {
        statusCb("Required GATT characteristics not found. Re-pair.");
        return cj;
    }

    AttachCommandResponseLogger(cj.cmdRespBasicChar, "basic 0x001A");
    AttachCommandResponseLogger(cj.cmdRespExtChar, "extended 0x001E");
    try {
        cj.device.RequestPreferredConnectionParameters(BluetoothLEPreferredConnectionParameters::ThroughputOptimized());
    } catch(...) {}
    outSuccess = true;
    statusCb("Connected!");
    return cj;
}

static void InitViGEm() {
    if (g_vigem) return;
    g_vigem = vigem_alloc();
    if (!g_vigem || !VIGEM_SUCCESS(vigem_connect(g_vigem))) {
        AppLog("[ERROR] ViGEm failed to connect");
    }
}
static PVIGEM_TARGET AddDS4() {
    auto t = vigem_target_ds4_alloc();
    auto err = vigem_target_add(g_vigem, t);
    if (!VIGEM_SUCCESS(err))
        AppLog("[ERROR] vigem_target_add failed");
    return t;
}

static void ApplyBtnMapping(DS4_REPORT_EX& r, ButtonMapping m) {
    switch(m) {
        case ButtonMapping::L3:       r.Report.wButtons |= DS4_BUTTON_THUMB_LEFT;    break;
        case ButtonMapping::R3:       r.Report.wButtons |= DS4_BUTTON_THUMB_RIGHT;   break;
        case ButtonMapping::L1:       r.Report.wButtons |= DS4_BUTTON_SHOULDER_LEFT; break;
        case ButtonMapping::R1:       r.Report.wButtons |= DS4_BUTTON_SHOULDER_RIGHT;break;
        case ButtonMapping::L2:       r.Report.bTriggerL = 255;                      break;
        case ButtonMapping::R2:       r.Report.bTriggerR = 255;                      break;
        case ButtonMapping::CROSS:    r.Report.wButtons |= DS4_BUTTON_CROSS;         break;
        case ButtonMapping::CIRCLE:   r.Report.wButtons |= DS4_BUTTON_CIRCLE;        break;
        case ButtonMapping::SQUARE:   r.Report.wButtons |= DS4_BUTTON_SQUARE;        break;
        case ButtonMapping::TRIANGLE: r.Report.wButtons |= DS4_BUTTON_TRIANGLE;      break;
        case ButtonMapping::SHARE:    r.Report.wButtons |= DS4_BUTTON_SHARE;         break;
        case ButtonMapping::OPTIONS:  r.Report.wButtons |= DS4_BUTTON_OPTIONS;       break;
        case ButtonMapping::DPAD_UP:    DS4_SET_DPAD(reinterpret_cast<PDS4_REPORT>(&r.Report), DS4_BUTTON_DPAD_NORTH); break;
        case ButtonMapping::DPAD_DOWN:  DS4_SET_DPAD(reinterpret_cast<PDS4_REPORT>(&r.Report), DS4_BUTTON_DPAD_SOUTH); break;
        case ButtonMapping::DPAD_LEFT:  DS4_SET_DPAD(reinterpret_cast<PDS4_REPORT>(&r.Report), DS4_BUTTON_DPAD_WEST);  break;
        case ButtonMapping::DPAD_RIGHT: DS4_SET_DPAD(reinterpret_cast<PDS4_REPORT>(&r.Report), DS4_BUTTON_DPAD_EAST);  break;
        default: break;
    }
}
static void ApplyGLGR(DS4_REPORT_EX& r, const std::vector<uint8_t>& buf) {
    if (buf.size() < 9 || g_proConfig.layouts.empty()) return;
    int li = g_proConfig.activeLayoutIndex;
    if (li < 0 || li >= (int)g_proConfig.layouts.size()) li = 0;
    auto& lay = g_proConfig.layouts[li];
    uint64_t st = 0;
    for (int i = 3; i <= 8; ++i) st = (st << 8) | buf[i];
    if (st & 0x000000000200ULL) ApplyBtnMapping(r, lay.glMapping);
    if (st & 0x000000000100ULL) ApplyBtnMapping(r, lay.grMapping);
}
static void HandleSpecialProButtons(const std::vector<uint8_t>& buf) {
    if (buf.size() < 9) return;
    uint64_t st = 0;
    for (int i = 3; i <= 8; ++i) st = (st << 8) | buf[i];
    constexpr uint64_t SCREENSHOT = 0x000020000000;
    constexpr uint64_t CBTN       = 0x000040000000;
    constexpr uint64_t GL         = 0x000000000200;
    constexpr uint64_t GR         = 0x000000000100;
    constexpr uint64_t ZL         = 0x000000800000;
    constexpr uint64_t ZR         = 0x008000000000;

    bool combo = (st & ZL) && (st & ZR) && (st & GL) && (st & GR);
    if (combo && !g_comboPressed) { g_openLayoutManager.store(true); g_comboPressed = true; }
    else if (!combo) g_comboPressed = false;

    bool screenshot = (st & SCREENSHOT) != 0;
    if (screenshot && !g_screenshotButtonPressed) { INPUT ip{}; ip.type=INPUT_KEYBOARD; ip.ki.wVk=VK_F12; SendInput(1,&ip,sizeof(ip)); g_screenshotButtonPressed=true; }
    else if (!screenshot && g_screenshotButtonPressed) { INPUT ip{}; ip.type=INPUT_KEYBOARD; ip.ki.wVk=VK_F12; ip.ki.dwFlags=KEYEVENTF_KEYUP; SendInput(1,&ip,sizeof(ip)); g_screenshotButtonPressed=false; }

    bool cBtn = (st & CBTN) != 0;
    if (cBtn && !g_cButtonPressed) {
        if (!g_proConfig.layouts.empty()) {
            g_proConfig.activeLayoutIndex = (g_proConfig.activeLayoutIndex + 1) % (int)g_proConfig.layouts.size();
            SaveProConfig();
            AppLog(std::string("Layout -> ") + g_proConfig.layouts[g_proConfig.activeLayoutIndex].name);
        }
        g_cButtonPressed = true;
    } else if (!cBtn) g_cButtonPressed = false;
}

static void SendKey(WORD vk, bool down) {
    INPUT ip{}; ip.type=INPUT_KEYBOARD; ip.ki.wVk=vk;
    ip.ki.dwFlags = down ? 0 : KEYEVENTF_KEYUP;
    SendInput(1, &ip, sizeof(ip));
}

struct CalibWizard {
    bool active   = false;
    int  step     = 0;
    bool isLeft   = true;
    int rawX = 2048, rawY = 2048;
    int centerX = 2048, centerY = 2048;
    int minX = 4095, maxX = 0, minY = 4095, maxY = 0;
    bool capturing = false;
    int  captureFrames = 0;
    std::vector<uint8_t> lastBuf;
};
static CalibWizard  g_calib;
static std::mutex   g_calibBufMutex;

static void ResetCalib(bool isLeft) {
    g_calib.active        = true;
    g_calib.step          = 0;
    g_calib.isLeft        = isLeft;
    g_calib.rawX          = 2048;
    g_calib.rawY          = 2048;
    g_calib.centerX       = 2048;
    g_calib.centerY       = 2048;
    g_calib.minX          = 4095;
    g_calib.maxX          = 0;
    g_calib.minY          = 4095;
    g_calib.maxY          = 0;
    g_calib.capturing     = false;
    g_calib.captureFrames = 0;
    {
        std::lock_guard<std::mutex> lk(g_calibBufMutex);
        g_calib.lastBuf.clear();
    }
}

static void FeedCalibBuffer(const std::vector<uint8_t>& buf, bool isLeftSide) {
    if (!g_calib.active) return;
    if (g_calib.isLeft != isLeftSide) return;
    std::lock_guard<std::mutex> lk(g_calibBufMutex);
    g_calib.lastBuf = buf;
}

static void UpdateCalibLiveValues() {
    std::lock_guard<std::mutex> lk(g_calibBufMutex);
    if (g_calib.lastBuf.size() < 16) return;
    ExtractRawStick(g_calib.lastBuf, g_calib.isLeft, g_calib.rawX, g_calib.rawY);
    if (g_calib.capturing) {
        g_calib.minX = std::min(g_calib.minX, g_calib.rawX);
        g_calib.maxX = std::max(g_calib.maxX, g_calib.rawX);
        g_calib.minY = std::min(g_calib.minY, g_calib.rawY);
        g_calib.maxY = std::max(g_calib.maxY, g_calib.rawY);
        ++g_calib.captureFrames;
    }
}

static void AttachSingleJoyConHandler(SingleJoyConPlayer& player, GyroMode gyroMode, uint8_t dsuSlot) {
    player.joycon.inputChar.ValueChanged(
        [&player, gyroMode, dsuSlot]
        (GattCharacteristic const&, GattValueChangedEventArgs const& args)
    {
        if (g_shuttingDown.load()) return;
        const auto now = SteadyClock::now();
        auto rdr = DataReader::FromBuffer(args.CharacteristicValue());
        std::vector<uint8_t> buf(rdr.UnconsumedBufferLength());
        rdr.ReadBytes(buf);

        FeedCalibBuffer(buf, player.side == JoyConSide::Left);

        const double bleDelta = MsBetween(player.latency.lastBleTime, now);
        player.latency.lastBleTime = now;

        if (player.side == JoyConSide::Right) {
            uint32_t btnState = ExtractButtonState(buf);
            bool chatPressed = (btnState & 0x000040) != 0;
            if (chatPressed && !player.wasChatPressed) {
                player.mouseMode = (player.mouseMode + 1) % 4;
                const char* names[] = {"OFF","FAST","NORMAL","SLOW"};
                uint8_t leds[] = {0x01, 0x02, 0x04, 0x08};
                AppLog(std::string("Mouse mode: ") + names[player.mouseMode]);
                SetPlayerLEDs(player.joycon.writeChar, leds[player.mouseMode]);
                EmitSound(player.joycon.writeChar);
            }
            player.wasChatPressed = chatPressed;

            if (player.mouseMode > 0) {
                auto [rx, ry] = GetRawOpticalMouse(buf);
                if (player.firstOpticalRead) { player.lastOpticalX=rx; player.lastOpticalY=ry; player.firstOpticalRead=false; }
                else {
                    int16_t dx=rx-player.lastOpticalX, dy=ry-player.lastOpticalY;
                    player.lastOpticalX=rx; player.lastOpticalY=ry;
                    if (dx||dy) {
                        float s = player.mouseMode==1?1.f:player.mouseMode==2?0.6f:0.3f;
                        INPUT ip{}; ip.type=INPUT_MOUSE; ip.mi.dx=(int)(dx*s); ip.mi.dy=(int)(dy*s); ip.mi.dwFlags=MOUSEEVENTF_MOVE;
                        SendInput(1,&ip,sizeof(ip));
                    }
                }
                bool R=(btnState&0x004000)!=0, ZR=(btnState&0x008000)!=0, ST=(btnState&0x000004)!=0;
                auto mkMouse=[](DWORD f){INPUT ip{}; ip.type=INPUT_MOUSE; ip.mi.dwFlags=f; SendInput(1,&ip,sizeof(ip));};
                if (R&&!player.leftBtnPressed)   mkMouse(MOUSEEVENTF_LEFTDOWN);
                if (!R&&player.leftBtnPressed)   mkMouse(MOUSEEVENTF_LEFTUP);
                player.leftBtnPressed=R;
                if (ZR&&!player.rightBtnPressed) mkMouse(MOUSEEVENTF_RIGHTDOWN);
                if (!ZR&&player.rightBtnPressed) mkMouse(MOUSEEVENTF_RIGHTUP);
                player.rightBtnPressed=ZR;
                if (ST&&!player.middleBtnPressed) mkMouse(MOUSEEVENTF_MIDDLEDOWN);
                if (!ST&&player.middleBtnPressed) mkMouse(MOUSEEVENTF_MIDDLEUP);
                player.middleBtnPressed=ST;

                auto sd = DecodeJoystick(buf, player.side, player.orientation);
                const int SZ=4000, BT=28000;
                if (abs(sd.y)>SZ) {
                    float inten=(abs(sd.y)-SZ)/(32767.f-SZ);
                    float spd=inten*40.f;
                    player.scrollAccumulator += sd.y>0 ? -spd : spd;
                    if (abs(player.scrollAccumulator)>=120.f) {
                        int clicks=(int)(player.scrollAccumulator/120.f);
                        player.scrollAccumulator-=clicks*120.f;
                        INPUT ip{}; ip.type=INPUT_MOUSE; ip.mi.mouseData=clicks*120; ip.mi.dwFlags=MOUSEEVENTF_WHEEL;
                        SendInput(1,&ip,sizeof(ip));
                    }
                } else player.scrollAccumulator=0.f;
                auto mkX=[](DWORD xb, DWORD f){INPUT ip{}; ip.type=INPUT_MOUSE; ip.mi.mouseData=xb; ip.mi.dwFlags=f; SendInput(1,&ip,sizeof(ip));};
                if (sd.x<-BT&&!player.mb4Pressed) { mkX(XBUTTON1,MOUSEEVENTF_XDOWN); mkX(XBUTTON1,MOUSEEVENTF_XUP); player.mb4Pressed=true; }
                else if (sd.x>=-BT) player.mb4Pressed=false;
                if (sd.x>BT&&!player.mb5Pressed) { mkX(XBUTTON2,MOUSEEVENTF_XDOWN); mkX(XBUTTON2,MOUSEEVENTF_XUP); player.mb5Pressed=true; }
                else if (sd.x<=BT) player.mb5Pressed=false;
                buf[4]&=~0x40; buf[4]&=~0x80; buf[5]&=~0x04;
                if (buf.size()>=16){buf[13]=0x00;buf[14]=0x08;buf[15]=0x80;}
            } else player.firstOpticalRead=true;
        }

        if (!ShouldEmit(g_opts.updatePolicy, player.latency.lastEmitTime, SteadyClock::now())) return;
        const auto ds = SteadyClock::now();
        DS4_REPORT_EX report = GenerateDS4Report(buf, player.side, player.orientation);
        if (gyroMode==GyroMode::DsuUdp && g_dsuServer.IsRunning()) {
            g_dsuServer.UpdateController(dsuSlot, report); 
        }
        if (g_shuttingDown.load() || !g_vigem || !player.ds4Controller) return;
        vigem_target_ds4_update_ex(g_vigem, player.ds4Controller, report);
        const auto vc = SteadyClock::now();
        g_latencyLogger.Record(g_opts.updatePolicy, CtrlTypeName(1), ++player.latency.eventIndex,
                               bleDelta, 0.0, -1.0, UsBetween(ds,vc), UsBetween(now,vc));
    });
}

static ID3D11Device*           g_pd3dDevice       = nullptr;
static ID3D11DeviceContext*    g_pd3dDeviceContext = nullptr;
static IDXGISwapChain*         g_pSwapChain        = nullptr;
static ID3D11RenderTargetView* g_mainRTV           = nullptr;
static HWND                    g_hwnd              = nullptr;

static void CreateRTV() {
    ID3D11Texture2D* bb = nullptr;
    g_pSwapChain->GetBuffer(0, IID_PPV_ARGS(&bb));
    g_pd3dDevice->CreateRenderTargetView(bb, nullptr, &g_mainRTV);
    bb->Release();
}
static void CleanupRTV() { if (g_mainRTV){g_mainRTV->Release();g_mainRTV=nullptr;} }
static bool CreateDX11(HWND hwnd) {
    DXGI_SWAP_CHAIN_DESC sd{};
    sd.BufferCount=2; sd.BufferDesc.Width=0; sd.BufferDesc.Height=0;
    sd.BufferDesc.Format=DXGI_FORMAT_R8G8B8A8_UNORM;
    sd.BufferDesc.RefreshRate.Numerator=60; sd.BufferDesc.RefreshRate.Denominator=1;
    sd.Flags=DXGI_SWAP_CHAIN_FLAG_ALLOW_MODE_SWITCH;
    sd.BufferUsage=DXGI_USAGE_RENDER_TARGET_OUTPUT; sd.OutputWindow=hwnd;
    sd.SampleDesc.Count=1; sd.Windowed=TRUE; sd.SwapEffect=DXGI_SWAP_EFFECT_DISCARD;
    D3D_FEATURE_LEVEL fl;
    if (FAILED(D3D11CreateDeviceAndSwapChain(nullptr,D3D_DRIVER_TYPE_HARDWARE,nullptr,0,nullptr,0,
        D3D11_SDK_VERSION,&sd,&g_pSwapChain,&g_pd3dDevice,&fl,&g_pd3dDeviceContext))) return false;
    CreateRTV(); return true;
}
static void CleanupDX11() {
    CleanupRTV();
    if (g_pSwapChain)        { g_pSwapChain->Release();        g_pSwapChain=nullptr; }
    if (g_pd3dDeviceContext) { g_pd3dDeviceContext->Release();  g_pd3dDeviceContext=nullptr; }
    if (g_pd3dDevice)        { g_pd3dDevice->Release();         g_pd3dDevice=nullptr; }
}

static void RequestImmediateExit() {
    g_shuttingDown.store(true);
    ExitProcess(0);
}

static LRESULT WINAPI WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    if (ImGui_ImplWin32_WndProcHandler(hWnd, msg, wParam, lParam)) return true;
    switch(msg) {
        case WM_CLOSE:
            RequestImmediateExit();
            return 0;
        case WM_SIZE:
            if (g_pd3dDevice && wParam != SIZE_MINIMIZED) {
                CleanupRTV();
                g_pSwapChain->ResizeBuffers(0,(UINT)LOWORD(lParam),(UINT)HIWORD(lParam),DXGI_FORMAT_UNKNOWN,0);
                CreateRTV();
            }
            return 0;
        case WM_SYSCOMMAND:
            if ((wParam&0xfff0)==SC_KEYMENU) return 0;
            break;
        case WM_DESTROY:
            g_shuttingDown.store(true);
            g_hwnd = nullptr;
            PostQuitMessage(0);
            return 0;
    }
    return DefWindowProcW(hWnd, msg, wParam, lParam);
}

static void HelpMarker(const char* txt) {
    ImGui::TextDisabled("(?)");
    if (ImGui::IsItemHovered()) { ImGui::BeginTooltip(); ImGui::TextUnformatted(txt); ImGui::EndTooltip(); }
}

static void DrawPlayerConfigRow(int i, PlayerConfig& cfg) {
    ImGui::PushID(i);
    ImGui::TableNextRow();

    ImGui::TableSetColumnIndex(0);
    ImGui::Text("Player %d", i+1);

    ImGui::TableSetColumnIndex(1);
    const char* ctypes[] = {"Single JoyCon","Dual JoyCon","Pro Controller","NSO GC"};
    int ct = (int)cfg.controllerType - 1;
    ImGui::SetNextItemWidth(-1);
    if (ImGui::Combo("##type", &ct, ctypes, 4))
        cfg.controllerType = (ControllerType)(ct+1);

    ImGui::TableSetColumnIndex(2);
    if (cfg.controllerType == SingleJoyCon) {
        const char* sides[] = {"Left","Right"};
        int s = (cfg.joyconSide == JoyConSide::Left) ? 0 : 1;
        ImGui::SetNextItemWidth(-1);
        if (ImGui::Combo("##side", &s, sides, 2))
            cfg.joyconSide = s==0 ? JoyConSide::Left : JoyConSide::Right;
    } else ImGui::TextDisabled("—");

    ImGui::TableSetColumnIndex(3);
    if (cfg.controllerType == SingleJoyCon) {
        const char* orients[] = {"Upright","Sideways"};
        int o = (cfg.joyconOrientation == JoyConOrientation::Upright) ? 0 : 1;
        ImGui::SetNextItemWidth(-1);
        if (ImGui::Combo("##orient", &o, orients, 2))
            cfg.joyconOrientation = o==0 ? JoyConOrientation::Upright : JoyConOrientation::Sideways;
    } else ImGui::TextDisabled("—");

    ImGui::TableSetColumnIndex(4);
    if (cfg.controllerType == DualJoyCon) {
        const char* gs[] = {"Both","Left","Right"};
        int g = (cfg.gyroSource==GyroSource::Both)?0:(cfg.gyroSource==GyroSource::Left)?1:2;
        ImGui::SetNextItemWidth(-1);
        if (ImGui::Combo("##gyrosrc", &g, gs, 3))
            cfg.gyroSource = g==0?GyroSource::Both:g==1?GyroSource::Left:GyroSource::Right;
    } else ImGui::TextDisabled("—");

    ImGui::TableSetColumnIndex(5);
    if (cfg.controllerType != NSOGCController) {
        const char* gm[] = {"Raw","DSU UDP"};
        int gv = (cfg.gyroMode==GyroMode::Raw)?0:1;
        ImGui::SetNextItemWidth(-1);
        if (ImGui::Combo("##gyromode", &gv, gm, 2)) {
            cfg.gyroMode = gv==0?GyroMode::Raw:GyroMode::DsuUdp;
        }
    } else ImGui::TextDisabled("—");

    ImGui::PopID();
}

static void DrawSetupScreen() {
    ImGuiIO& io = ImGui::GetIO();
    ImGui::SetNextWindowPos({0,0}); ImGui::SetNextWindowSize(io.DisplaySize);
    ImGui::Begin("##setup", nullptr,
        ImGuiWindowFlags_NoTitleBar|ImGuiWindowFlags_NoResize|ImGuiWindowFlags_NoMove|ImGuiWindowFlags_NoCollapse);

    ImGui::Spacing();
    ImGui::SetCursorPosX((io.DisplaySize.x - ImGui::CalcTextSize("joycon2cpp").x) * 0.5f);
    ImGui::TextColored({0.4f,0.8f,1.f,1.f}, "joycon2cpp");
    ImGui::Separator(); ImGui::Spacing();

    if (ImGui::CollapsingHeader("Players", ImGuiTreeNodeFlags_DefaultOpen)) {
        ImGui::Spacing();
        if (ImGui::BeginTable("players", 6,
            ImGuiTableFlags_Borders|ImGuiTableFlags_RowBg|ImGuiTableFlags_SizingStretchProp)) {
            ImGui::TableSetupColumn("Player",       ImGuiTableColumnFlags_WidthFixed, 60);
            ImGui::TableSetupColumn("Type",         ImGuiTableColumnFlags_WidthStretch, 1.4f);
            ImGui::TableSetupColumn("Side",         ImGuiTableColumnFlags_WidthStretch, 0.8f);
            ImGui::TableSetupColumn("Orientation",  ImGuiTableColumnFlags_WidthStretch, 0.9f);
            ImGui::TableSetupColumn("Gyro Source",  ImGuiTableColumnFlags_WidthStretch, 0.9f);
            ImGui::TableSetupColumn("Gyro Output",  ImGuiTableColumnFlags_WidthStretch, 1.2f);
            ImGui::TableHeadersRow();

            for (int i = 0; i < (int)g_playerConfigs.size(); ++i)
                DrawPlayerConfigRow(i, g_playerConfigs[i]);

            ImGui::EndTable();
        }
        ImGui::Spacing();
        if (ImGui::Button("+ Add Player") && g_playerConfigs.size() < 4)
            g_playerConfigs.push_back({});
        ImGui::SameLine();
        if (ImGui::Button("- Remove Last") && g_playerConfigs.size() > 1)
            g_playerConfigs.pop_back();
        ImGui::Spacing();
    }

    if (ImGui::CollapsingHeader("Settings")) {
        ImGui::Spacing();
        ImGui::Indent(10);

        const char* policies[] = {"Low Latency (immediate)","Balanced 120Hz","Legacy 60Hz"};
        int pol = (int)g_opts.updatePolicy;
        ImGui::SetNextItemWidth(220);
        if (ImGui::Combo("Update Policy", &pol, policies, 3))
            g_opts.updatePolicy = (UpdatePolicy)pol;
        ImGui::SameLine(); HelpMarker("Low Latency forwards every BLE packet immediately.\nBalanced/Legacy cap the output rate.");

        ImGui::Checkbox("Record latency metrics to CSV", &g_opts.latencyMetrics);
        if (g_opts.latencyMetrics) {
            ImGui::SetNextItemWidth(300);
            ImGui::InputText("CSV path", g_opts.latencyCsvPath, sizeof(g_opts.latencyCsvPath));
        }

        ImGui::Unindent(10);
        ImGui::Spacing();
    }

    if (ImGui::CollapsingHeader("Stick Calibration")) {
        ImGui::Spacing();
        ImGui::Indent(10);
        ImGui::TextWrapped("Calibrate after connecting. Connect your controllers first, then open this panel in-session.");

        const auto& cal = GetActiveCalibration();
        ImGui::Spacing();
        if (ImGui::BeginTable("calvals", 4, ImGuiTableFlags_Borders|ImGuiTableFlags_SizingStretchSame)) {
            ImGui::TableSetupColumn("Stick"); ImGui::TableSetupColumn("Center");
            ImGui::TableSetupColumn("Range X"); ImGui::TableSetupColumn("Range Y");
            ImGui::TableHeadersRow();
            ImGui::TableNextRow();
            ImGui::TableSetColumnIndex(0); ImGui::Text("Left");
            ImGui::TableSetColumnIndex(1); ImGui::Text("%d, %d", cal.leftStick.centerX, cal.leftStick.centerY);
            ImGui::TableSetColumnIndex(2); ImGui::Text("%d–%d", cal.leftStick.minX, cal.leftStick.maxX);
            ImGui::TableSetColumnIndex(3); ImGui::Text("%d–%d", cal.leftStick.minY, cal.leftStick.maxY);
            ImGui::TableNextRow();
            ImGui::TableSetColumnIndex(0); ImGui::Text("Right");
            ImGui::TableSetColumnIndex(1); ImGui::Text("%d, %d", cal.rightStick.centerX, cal.rightStick.centerY);
            ImGui::TableSetColumnIndex(2); ImGui::Text("%d–%d", cal.rightStick.minX, cal.rightStick.maxX);
            ImGui::TableSetColumnIndex(3); ImGui::Text("%d–%d", cal.rightStick.minY, cal.rightStick.maxY);
            ImGui::EndTable();
        }

        ImGui::Spacing();
        if (ImGui::Button("Reset to defaults")) {
            CalibrationProfile def;
            def.name = "Default";
            while (GetCalibrationProfiles().size() > 1)
                DeleteCalibrationProfile(1);
            DeleteCalibrationProfile(0);
            AddCalibrationProfile(def);
            SetActiveCalibrationIndex(0);
            SaveCalibrationProfiles("calibration.json");
        }
        ImGui::Unindent(10);
        ImGui::Spacing();
    }

    bool hasProController = false;
    for (auto& pc : g_playerConfigs) if (pc.controllerType == ProController) { hasProController = true; break; }
    if (hasProController && ImGui::CollapsingHeader("Pro Controller GL/GR Layouts")) {
        ImGui::Spacing();
        ImGui::Indent(10);
        ImGui::Text("Active layout: %s", g_proConfig.layouts[g_proConfig.activeLayoutIndex].name);
        ImGui::Spacing();
        for (int i = 0; i < (int)g_proConfig.layouts.size(); ++i) {
            ImGui::PushID(i);
            auto& lay = g_proConfig.layouts[i];
            bool isActive = (i == g_proConfig.activeLayoutIndex);
            if (isActive) ImGui::PushStyleColor(ImGuiCol_Header, ImGui::GetStyleColorVec4(ImGuiCol_HeaderActive));
            bool open = ImGui::TreeNodeEx(lay.name, ImGuiTreeNodeFlags_Framed);
            if (isActive) ImGui::PopStyleColor();
            if (open) {
                ImGui::InputText("Name##n", lay.name, 64);
                int gl = (int)lay.glMapping, gr = (int)lay.grMapping;
                ImGui::SetNextItemWidth(160); ImGui::Combo("GL##gl", &gl, BtnMapNames, 17); lay.glMapping=(ButtonMapping)gl;
                ImGui::SameLine();
                ImGui::SetNextItemWidth(160); ImGui::Combo("GR##gr", &gr, BtnMapNames, 17); lay.grMapping=(ButtonMapping)gr;
                ImGui::Spacing();
                if (!isActive && ImGui::Button("Set Active")) { g_proConfig.activeLayoutIndex=i; SaveProConfig(); }
                ImGui::SameLine();
                if (g_proConfig.layouts.size()>1 && ImGui::Button("Delete")) {
                    g_proConfig.layouts.erase(g_proConfig.layouts.begin()+i);
                    if (g_proConfig.activeLayoutIndex>=(int)g_proConfig.layouts.size())
                        g_proConfig.activeLayoutIndex=(int)g_proConfig.layouts.size()-1;
                    SaveProConfig(); ImGui::TreePop(); ImGui::PopID(); break;
                }
                SaveProConfig();
                ImGui::TreePop();
            }
            ImGui::PopID();
        }
        if (ImGui::Button("+ New Layout")) {
            GLGRLayout nl; snprintf(nl.name,64,"Layout %d",(int)g_proConfig.layouts.size()+1);
            g_proConfig.layouts.push_back(nl); SaveProConfig();
        }
        ImGui::Unindent(10);
        ImGui::Spacing();
    }

    ImGui::Spacing(); ImGui::Separator(); ImGui::Spacing();

    float btnW = 180;
    ImGui::SetCursorPosX((io.DisplaySize.x - btnW) * 0.5f);
    ImGui::PushStyleColor(ImGuiCol_Button,        {0.2f,0.6f,0.2f,1.f});
    ImGui::PushStyleColor(ImGuiCol_ButtonHovered, {0.3f,0.75f,0.3f,1.f});
    ImGui::PushStyleColor(ImGuiCol_ButtonActive,  {0.15f,0.5f,0.15f,1.f});
    bool doConnect = ImGui::Button("Connect Controllers", {btnW, 36});
    ImGui::PopStyleColor(3);

    if (doConnect && !g_playerConfigs.empty()) {
        g_connectionTasks.clear();
        for (int i = 0; i < (int)g_playerConfigs.size(); ++i) {
            auto& pc = g_playerConfigs[i];
            if (pc.controllerType == SingleJoyCon) {
                std::string side = (pc.joyconSide==JoyConSide::Left)?"Left":"Right";
                g_connectionTasks.push_back({"Player " + std::to_string(i+1) + " — " + side + " JoyCon"});
            } else if (pc.controllerType == DualJoyCon) {
                g_connectionTasks.push_back({"Player " + std::to_string(i+1) + " — RIGHT JoyCon"});
                g_connectionTasks.push_back({"Player " + std::to_string(i+1) + " — LEFT JoyCon"});
            } else if (pc.controllerType == ProController) {
                g_connectionTasks.push_back({"Player " + std::to_string(i+1) + " — Pro Controller"});
            } else {
                g_connectionTasks.push_back({"Player " + std::to_string(i+1) + " — NSO GC Controller"});
            }
        }
        g_connectionTaskIndex = 0;
        g_connectionDone      = false;
        g_connectionError     = "";
        g_screen = AppScreen::Connecting;

        InitViGEm();
        bool needsDsu = false;
        for (auto& pc : g_playerConfigs) if (pc.gyroMode==GyroMode::DsuUdp) { needsDsu=true; break; }
        if (needsDsu) g_dsuServer.Start();
        if (g_opts.latencyMetrics)
            g_latencyLogger.Start(g_opts.latencyCsvPath);

        std::thread([configs = g_playerConfigs]() mutable {
            int taskIdx = 0;

            g_singlePlayers.clear();
            g_dualPlayers.clear();
            g_proPlayers.clear();

            for (auto* p : g_singleRumbleCtxs) delete p;
            for (auto* p : g_dualRumbleCtxs)   delete p;
            for (auto* p : g_proRumbleCtxs)     delete p;
            g_singleRumbleCtxs.clear();
            g_dualRumbleCtxs.clear();
            g_proRumbleCtxs.clear();

            int dsuSlot = 0;
            for (int pi = 0; pi < (int)configs.size(); ++pi) {
                auto& pc = configs[pi];

                if (pc.controllerType == SingleJoyCon) {
                    g_connectionTasks[taskIdx].statusMsg = "Scanning...";
                    bool ok = false;
                    auto cj = BleConnect([&](const std::string& s){ g_connectionTasks[taskIdx].statusMsg=s; }, ok, pc.joyconSide);
                    if (!ok) { g_connectionError="Failed: "+g_connectionTasks[taskIdx].label; g_connectionDone=true; return; }
                    g_connectionTasks[taskIdx].done=true; g_connectionTasks[taskIdx].success=true;
                    if (cj.rumbleChar) { SendJoyCon2OfficialInit(cj.rumbleChar); }
                    auto t = AddDS4();
                    g_singlePlayers.push_back({ cj, t, pc.joyconSide, pc.joyconOrientation });
                    auto& player = g_singlePlayers.back();
                    if (pc.gyroMode==GyroMode::DsuUdp && g_dsuServer.IsRunning()) g_dsuServer.SetControllerConnected(dsuSlot);
                    AttachSingleJoyConHandler(player, pc.gyroMode, dsuSlot);
                    cj.inputChar.WriteClientCharacteristicConfigurationDescriptorAsync(GattClientCharacteristicConfigurationDescriptorValue::Notify).get();

                    auto* rctx = new SingleRumbleCtx{ cj.vibrationChar, t };
                    g_singleRumbleCtxs.push_back(rctx);
                    StartSingleRumbleThread(rctx);

                    AppLog("Player " + std::to_string(pi+1) + " Single JoyCon connected");
                    ++taskIdx; ++dsuSlot;

                } else if (pc.controllerType == DualJoyCon) {
                    g_connectionTasks[taskIdx].statusMsg = "Scanning...";
                    bool okR = false;
                    auto rjc = BleConnect([&](const std::string& s){ g_connectionTasks[taskIdx].statusMsg=s; }, okR, JoyConSide::Right);
                    if (!okR) { g_connectionError="Failed: "+g_connectionTasks[taskIdx].label; g_connectionDone=true; return; }
                    g_connectionTasks[taskIdx].done=true; g_connectionTasks[taskIdx].success=true;
                    if (rjc.rumbleChar) { SendJoyCon2OfficialInit(rjc.rumbleChar); }
                    ++taskIdx;

                    g_connectionTasks[taskIdx].statusMsg = "Scanning...";
                    bool okL = false;
                    auto ljc = BleConnect([&](const std::string& s){ g_connectionTasks[taskIdx].statusMsg=s; }, okL, JoyConSide::Left);
                    if (!okL) { g_connectionError="Failed: "+g_connectionTasks[taskIdx].label; g_connectionDone=true; return; }
                    g_connectionTasks[taskIdx].done=true; g_connectionTasks[taskIdx].success=true;
                    if (ljc.rumbleChar) { SendJoyCon2OfficialInit(ljc.rumbleChar); }
                    ++taskIdx;

                    auto dp = std::make_unique<DualJoyConPlayer>();
                    dp->leftJoyCon=ljc; dp->rightJoyCon=rjc;
                    dp->gyroSource=pc.gyroSource; dp->gyroMode=pc.gyroMode;
                    dp->dsuSlot=dsuSlot;
                    dp->ds4Controller=AddDS4(); dp->running.store(true);
                    dp->sharedState=std::make_shared<DualJoyConSharedState>();
                    if (dp->gyroMode==GyroMode::DsuUdp&&g_dsuServer.IsRunning()) g_dsuServer.SetControllerConnected(dsuSlot);
                    auto ss=dp->sharedState;
                    ljc.inputChar.ValueChanged([ss](GattCharacteristic const&, GattValueChangedEventArgs const& a){
                        if (g_shuttingDown.load()) return;
                        auto now=SteadyClock::now(); auto rdr=DataReader::FromBuffer(a.CharacteristicValue());
                        std::vector<uint8_t> buf(rdr.UnconsumedBufferLength()); rdr.ReadBytes(buf);
                        FeedCalibBuffer(buf, true);
                        std::lock_guard<std::mutex> lk(ss->mutex);
                        ss->left.bleDeltaMs=MsBetween(ss->lastLeftBleTime,now); ss->lastLeftBleTime=now;
                        ss->left.buffer=std::move(buf); ss->left.receivedAt=now; ss->left.sequence=++ss->sequence;
                        ss->cv.notify_one();
                    });
                    ljc.inputChar.WriteClientCharacteristicConfigurationDescriptorAsync(GattClientCharacteristicConfigurationDescriptorValue::Notify).get();
                    rjc.inputChar.ValueChanged([ss](GattCharacteristic const&, GattValueChangedEventArgs const& a){
                        if (g_shuttingDown.load()) return;
                        auto now=SteadyClock::now(); auto rdr=DataReader::FromBuffer(a.CharacteristicValue());
                        std::vector<uint8_t> buf(rdr.UnconsumedBufferLength()); rdr.ReadBytes(buf);
                        FeedCalibBuffer(buf, false);
                        std::lock_guard<std::mutex> lk(ss->mutex);
                        ss->right.bleDeltaMs=MsBetween(ss->lastRightBleTime,now); ss->lastRightBleTime=now;
                        ss->right.buffer=std::move(buf); ss->right.receivedAt=now; ss->right.sequence=++ss->sequence;
                        ss->cv.notify_one();
                    });
                    rjc.inputChar.WriteClientCharacteristicConfigurationDescriptorAsync(GattClientCharacteristicConfigurationDescriptorValue::Notify).get();
                    dp->updateThread=std::thread([dpptr=dp.get(),ss](){
                        uint64_t lastSeq=0;
                        while (dpptr->running.load(std::memory_order_acquire) && !g_shuttingDown.load()) {
                            TimedInputBuffer ls,rs;
                            {
                                std::unique_lock<std::mutex> lk(ss->mutex);
                                ss->cv.wait_for(lk,std::chrono::milliseconds(1),[&]{ return !dpptr->running.load()||ss->sequence!=lastSeq; });
                                if (!dpptr->running.load()) break;
                                if (ss->left.buffer.empty()||ss->right.buffer.empty()||ss->sequence==lastSeq) continue;
                                auto now=SteadyClock::now();
                                if (!ShouldEmit(g_opts.updatePolicy,ss->lastEmitTime,now)){lastSeq=ss->sequence;continue;}
                                ls=ss->left; rs=ss->right; lastSeq=ss->sequence;
                            }
                            auto report=GenerateDualJoyConDS4Report(ls.buffer,rs.buffer,dpptr->gyroSource);
                            if (dpptr->gyroMode==GyroMode::DsuUdp&&g_dsuServer.IsRunning()) {
                                g_dsuServer.UpdateController(dpptr->dsuSlot,report);
                            }
                            if (g_shuttingDown.load() || !g_vigem || !dpptr->ds4Controller) break;
                            vigem_target_ds4_update_ex(g_vigem,dpptr->ds4Controller,report);
                        }
                    });

                    auto* rctx = new DualRumbleCtx{ ljc.vibrationChar, rjc.vibrationChar, dp->ds4Controller };
                    g_dualRumbleCtxs.push_back(rctx);
                    StartDualRumbleThread(rctx);

                    g_dualPlayers.push_back(std::move(dp));
                    AppLog("Player " + std::to_string(pi+1) + " Dual JoyCon connected");
                    ++dsuSlot;

                } else if (pc.controllerType == ProController) {
                    g_connectionTasks[taskIdx].statusMsg = "Scanning...";
                    bool ok = false;
                    auto cj = BleConnect([&](const std::string& s){ g_connectionTasks[taskIdx].statusMsg=s; }, ok, JoyConSide::Right);
                    if (!ok) { g_connectionError="Failed: "+g_connectionTasks[taskIdx].label; g_connectionDone=true; return; }
                    g_connectionTasks[taskIdx].done=true; g_connectionTasks[taskIdx].success=true;
                    if (cj.rumbleChar) { SendJoyCon2OfficialInit(cj.rumbleChar); }
                    auto tgt=AddDS4();
                    auto latPtr=std::make_shared<LatencyTracker>();
                    auto gm=pc.gyroMode; uint8_t ds=(uint8_t)dsuSlot;
                    cj.inputChar.ValueChanged([tgt,gm,ds,latPtr](GattCharacteristic const&, GattValueChangedEventArgs const& a) mutable {
                        if (g_shuttingDown.load()) return;
                        auto now=SteadyClock::now(); auto rdr=DataReader::FromBuffer(a.CharacteristicValue());
                        std::vector<uint8_t> buf(rdr.UnconsumedBufferLength()); rdr.ReadBytes(buf);
                        FeedCalibBuffer(buf, g_calib.isLeft);
                        double bd=MsBetween(latPtr->lastBleTime,now); latPtr->lastBleTime=now;
                        if (!ShouldEmit(g_opts.updatePolicy,latPtr->lastEmitTime,SteadyClock::now())) return;
                        DS4_REPORT_EX report=GenerateProControllerReport(buf);
                        ApplyGLGR(report,buf); HandleSpecialProButtons(buf);
                        if (gm==GyroMode::DsuUdp&&g_dsuServer.IsRunning()) {
                            ApplyGLGR(report,buf); g_dsuServer.UpdateController(ds,report);
                        }
                        if (g_shuttingDown.load() || !g_vigem || !tgt) return;
                        if (g_shuttingDown.load() || !g_vigem || !tgt) return;
                        vigem_target_ds4_update_ex(g_vigem,tgt,report);
                    });
                    cj.inputChar.WriteClientCharacteristicConfigurationDescriptorAsync(GattClientCharacteristicConfigurationDescriptorValue::Notify).get();
                    if (pc.gyroMode==GyroMode::DsuUdp&&g_dsuServer.IsRunning()) g_dsuServer.SetControllerConnected(dsuSlot);

                    auto* rctx = new SingleRumbleCtx{ cj.vibrationCharLeft, tgt, cj.vibrationCharRight };
                    g_proRumbleCtxs.push_back(rctx);
                    StartSingleRumbleThread(rctx);

                    g_proPlayers.push_back({cj,tgt});
                    AppLog("Player " + std::to_string(pi+1) + " Pro Controller connected");
                    ++taskIdx; ++dsuSlot;

                } else {
                    g_connectionTasks[taskIdx].statusMsg = "Scanning...";
                    bool ok = false;
                    auto cj = BleConnect([&](const std::string& s){ g_connectionTasks[taskIdx].statusMsg=s; }, ok);
                    if (!ok) { g_connectionError="Failed: "+g_connectionTasks[taskIdx].label; g_connectionDone=true; return; }
                    g_connectionTasks[taskIdx].done=true; g_connectionTasks[taskIdx].success=true;
                    if (cj.rumbleChar) { SendJoyCon2OfficialInit(cj.rumbleChar); }
                    auto tgt=AddDS4();
                    cj.inputChar.ValueChanged([tgt](GattCharacteristic const&, GattValueChangedEventArgs const& a) mutable {
                        if (g_shuttingDown.load()) return;
                        auto rdr=DataReader::FromBuffer(a.CharacteristicValue());
                        std::vector<uint8_t> buf(rdr.UnconsumedBufferLength()); rdr.ReadBytes(buf);
                        DS4_REPORT_EX report=GenerateNSOGCReport(buf);
                        if (g_shuttingDown.load() || !g_vigem || !tgt) return;
                        if (g_shuttingDown.load() || !g_vigem || !tgt) return;
                        vigem_target_ds4_update_ex(g_vigem,tgt,report);
                    });
                    cj.inputChar.WriteClientCharacteristicConfigurationDescriptorAsync(GattClientCharacteristicConfigurationDescriptorValue::Notify).get();
                    g_proPlayers.push_back({cj,tgt});
                    AppLog("Player " + std::to_string(pi+1) + " NSO GC connected");
                    ++taskIdx; ++dsuSlot;
                }
            }
            g_connectionDone = true;
        }).detach();
    }

    ImGui::End();
}

static void DrawConnectingScreen() {
    ImGuiIO& io = ImGui::GetIO();
    ImGui::SetNextWindowPos({0,0}); ImGui::SetNextWindowSize(io.DisplaySize);
    ImGui::Begin("##connecting", nullptr,
        ImGuiWindowFlags_NoTitleBar|ImGuiWindowFlags_NoResize|ImGuiWindowFlags_NoMove);

    ImGui::Spacing();
    ImGui::SetCursorPosX((io.DisplaySize.x - ImGui::CalcTextSize("Connecting...").x)*0.5f);
    ImGui::TextColored({0.4f,0.8f,1.f,1.f}, "Connecting...");
    ImGui::Spacing(); ImGui::Separator(); ImGui::Spacing();

    if (!g_connectionError.empty()) {
        ImGui::TextColored({1.f,0.3f,0.3f,1.f}, "Error: %s", g_connectionError.c_str());
        ImGui::Spacing();
        if (ImGui::Button("Back to Setup")) { g_connectionError=""; g_screen=AppScreen::Setup; }
    } else {
        for (auto& t : g_connectionTasks) {
            if (t.done) {
                ImGui::TextColored({0.3f,1.f,0.3f,1.f}, "[OK] %s", t.label.c_str());
            } else {
                static float spin = 0.f; spin += ImGui::GetIO().DeltaTime * 3.f;
                const char* spinners[] = {"|","/","-","\\"};
                ImGui::TextColored({1.f,0.8f,0.2f,1.f}, "[%s] %s — %s",
                    spinners[(int)(spin)%4], t.label.c_str(), t.statusMsg.c_str());
                break;
            }
        }

        ImGui::Spacing();
        if (g_connectionDone) {
            ImGui::TextColored({0.3f,1.f,0.3f,1.f}, "All controllers connected!");
            ImGui::Spacing();
            float bw=120;
            ImGui::SetCursorPosX((io.DisplaySize.x-bw)*0.5f);
            ImGui::PushStyleColor(ImGuiCol_Button,{0.2f,0.6f,0.2f,1.f});
            ImGui::PushStyleColor(ImGuiCol_ButtonHovered,{0.3f,0.75f,0.3f,1.f});
            ImGui::PushStyleColor(ImGuiCol_ButtonActive,{0.15f,0.5f,0.15f,1.f});
            if (ImGui::Button("Start!", {bw,36})) g_screen=AppScreen::Running;
            ImGui::PopStyleColor(3);
        }
    }
    ImGui::End();
}

static void DrawCalibWizard() {
    if (!g_calib.active) return;
    UpdateCalibLiveValues();

    ImGui::OpenPopup("Stick Calibration Wizard");
    ImVec2 center = ImGui::GetMainViewport()->GetCenter();
    ImGui::SetNextWindowPos(center, ImGuiCond_Always, {0.5f,0.5f});
    ImGui::SetNextWindowSize({480,360}, ImGuiCond_Always);

    if (ImGui::BeginPopupModal("Stick Calibration Wizard", &g_calib.active)) {
        const char* stickName = g_calib.isLeft ? "Left Stick" : "Right Stick";
        ImGui::TextColored({0.4f,0.8f,1.f,1.f}, "Calibrating: %s", stickName);
        ImGui::Separator(); ImGui::Spacing();

        ImGui::Text("Raw X: %4d   Raw Y: %4d", g_calib.rawX, g_calib.rawY);
        ImGui::Spacing();

        if (g_calib.step == 0) {
            ImGui::TextWrapped("This wizard will calibrate your stick's center and range.\n\n"
                               "Step 1: Release the stick and let it sit at rest. Then click Capture Center.");
            ImGui::Spacing();
            if (ImGui::Button("Capture Center", {160,32})) {
                g_calib.centerX = g_calib.rawX;
                g_calib.centerY = g_calib.rawY;
                g_calib.step = 1;
            }
        } else if (g_calib.step == 1) {
            ImGui::TextColored({0.3f,1.f,0.3f,1.f}, "Center captured: %d, %d", g_calib.centerX, g_calib.centerY);
            ImGui::Spacing();
            ImGui::TextWrapped("Step 2: Slowly rotate the stick all the way around in a full circle to capture the extents.\n"
                               "Click Start Capture, rotate, then click Done.");
            ImGui::Spacing();
            if (!g_calib.capturing) {
                if (ImGui::Button("Start Capturing Extents", {200,32})) {
                    g_calib.minX=4095; g_calib.maxX=0; g_calib.minY=4095; g_calib.maxY=0;
                    g_calib.captureFrames=0; g_calib.capturing=true;
                }
            } else {
                ImGui::TextColored({1.f,0.5f,0.1f,1.f}, "Capturing... rotate the stick fully!");
                ImGui::Text("Frames: %d   MinX:%d MaxX:%d MinY:%d MaxY:%d",
                    g_calib.captureFrames, g_calib.minX, g_calib.maxX, g_calib.minY, g_calib.maxY);
                if (ImGui::Button("Done Rotating", {160,32})) {
                    g_calib.capturing = false;
                    g_calib.step = 2;
                }
            }
        } else if (g_calib.step == 2) {
            ImGui::TextColored({0.3f,1.f,0.3f,1.f}, "Center: %d, %d", g_calib.centerX, g_calib.centerY);
            ImGui::TextColored({0.3f,1.f,0.3f,1.f}, "X Range: %d – %d", g_calib.minX, g_calib.maxX);
            ImGui::TextColored({0.3f,1.f,0.3f,1.f}, "Y Range: %d – %d", g_calib.minY, g_calib.maxY);
            ImGui::Spacing();
            ImGui::TextWrapped("Review the values above. Click Apply to save, or Redo to recapture.");
            ImGui::Spacing();
            if (ImGui::Button("Apply", {100,32})) {
                if (g_calib.minX >= g_calib.maxX || g_calib.minY >= g_calib.maxY
                    || g_calib.captureFrames < 10) {
                    AppLog("Calibration failed: rotate the stick fully before applying.");
                    g_calib.step = 1;
                } else {
                    CalibrationProfile updated = GetActiveCalibration();
                    auto& sc = g_calib.isLeft ? updated.leftStick : updated.rightStick;
                    sc.centerX = g_calib.centerX; sc.centerY = g_calib.centerY;
                    sc.minX = g_calib.minX;       sc.maxX = g_calib.maxX;
                    sc.minY = g_calib.minY;       sc.maxY = g_calib.maxY;
                    int idx = GetActiveCalibrationIndex();
                    DeleteCalibrationProfile(idx);
                    AddCalibrationProfile(updated);
                    SetActiveCalibrationIndex((int)GetCalibrationProfiles().size()-1);
                    SaveCalibrationProfiles("calibration.json");
                    AppLog(std::string("Calibration applied for ") + stickName);
                    g_calib.active = false;
                }
            }
            ImGui::SameLine();
            if (ImGui::Button("Redo", {100,32})) { g_calib.step=0; }
        }

        ImGui::Spacing(); ImGui::Separator();
        if (ImGui::Button("Cancel")) g_calib.active=false;
        ImGui::EndPopup();
    }
}

static void DrawLayoutManager() {
    if (!g_showLayoutManager) return;
    ImGui::OpenPopup("GL/GR Layout Manager");
    ImVec2 center = ImGui::GetMainViewport()->GetCenter();
    ImGui::SetNextWindowPos(center, ImGuiCond_Always, {0.5f,0.5f});
    ImGui::SetNextWindowSize({500,400}, ImGuiCond_Always);

    if (ImGui::BeginPopupModal("GL/GR Layout Manager", &g_showLayoutManager)) {
        ImGui::Text("Active: %s", g_proConfig.layouts[g_proConfig.activeLayoutIndex].name);
        ImGui::Separator(); ImGui::Spacing();

        for (int i = 0; i < (int)g_proConfig.layouts.size(); ++i) {
            ImGui::PushID(i);
            auto& lay = g_proConfig.layouts[i];
            bool isActive = (i == g_proConfig.activeLayoutIndex);
            if (isActive) ImGui::PushStyleColor(ImGuiCol_ChildBg, ImVec4(0.15f,0.3f,0.15f,1.f));
            ImGui::BeginChild(ImGui::GetID("##lay"), {-1,80}, true);
            ImGui::InputText("##name", lay.name, 64);
            ImGui::SameLine();
            if (isActive) ImGui::TextColored({0.3f,1.f,0.3f,1.f},"[ACTIVE]");
            int gl=(int)lay.glMapping, gr=(int)lay.grMapping;
            ImGui::SetNextItemWidth(140); ImGui::Combo("GL##g", &gl, BtnMapNames, 17); lay.glMapping=(ButtonMapping)gl;
            ImGui::SameLine();
            ImGui::SetNextItemWidth(140); ImGui::Combo("GR##r", &gr, BtnMapNames, 17); lay.grMapping=(ButtonMapping)gr;
            ImGui::SameLine();
            if (!isActive && ImGui::SmallButton("Set Active")) { g_proConfig.activeLayoutIndex=i; SaveProConfig(); }
            ImGui::SameLine();
            if (g_proConfig.layouts.size()>1 && ImGui::SmallButton("Del")) {
                g_proConfig.layouts.erase(g_proConfig.layouts.begin()+i);
                if (g_proConfig.activeLayoutIndex>=(int)g_proConfig.layouts.size())
                    g_proConfig.activeLayoutIndex=(int)g_proConfig.layouts.size()-1;
                SaveProConfig(); ImGui::EndChild(); if (isActive) ImGui::PopStyleColor(); ImGui::PopID(); break;
            }
            ImGui::EndChild();
            if (isActive) ImGui::PopStyleColor();
            ImGui::PopID();
        }

        ImGui::Spacing();
        if (ImGui::Button("+ New Layout")) {
            GLGRLayout nl; snprintf(nl.name,64,"Layout %d",(int)g_proConfig.layouts.size()+1);
            g_proConfig.layouts.push_back(nl); SaveProConfig();
        }
        ImGui::SameLine();
        if (ImGui::Button("Close")) { g_showLayoutManager=false; SaveProConfig(); }
        ImGui::EndPopup();
    }
}

static void DrawRunningScreen() {
    ImGuiIO& io = ImGui::GetIO();
    ImGui::SetNextWindowPos({0,0}); ImGui::SetNextWindowSize(io.DisplaySize);
    ImGui::Begin("##running", nullptr,
        ImGuiWindowFlags_NoTitleBar|ImGuiWindowFlags_NoResize|ImGuiWindowFlags_NoMove);

    ImGui::TextColored({0.4f,0.8f,1.f,1.f}, "joycon2cpp — Running");
    ImGui::SameLine(io.DisplaySize.x - 110);
    ImGui::PushStyleColor(ImGuiCol_Button,{0.6f,0.2f,0.2f,1.f});
    ImGui::PushStyleColor(ImGuiCol_ButtonHovered,{0.75f,0.3f,0.3f,1.f});
    ImGui::PushStyleColor(ImGuiCol_ButtonActive,{0.5f,0.15f,0.15f,1.f});
    if (ImGui::Button("Disconnect & Exit", {80,0})) {
        RequestImmediateExit();
    }
    ImGui::PopStyleColor(3);
    ImGui::Separator(); ImGui::Spacing();

    if (ImGui::CollapsingHeader("Connected Controllers", ImGuiTreeNodeFlags_DefaultOpen)) {
        ImGui::Indent(10);
        for (int i=0; i<(int)g_singlePlayers.size(); ++i) {
            auto& p=g_singlePlayers[i];
            ImGui::BulletText("Single JoyCon (%s, %s)",
                p.side==JoyConSide::Left?"Left":"Right",
                p.orientation==JoyConOrientation::Upright?"Upright":"Sideways");
        }
        for (int i=0; i<(int)g_dualPlayers.size(); ++i)
            ImGui::BulletText("Dual JoyCon pair");
        for (int i=0; i<(int)g_proPlayers.size(); ++i)
            ImGui::BulletText("Pro / NSO GC Controller");
        ImGui::Unindent(10); ImGui::Spacing();
    }

    if (ImGui::CollapsingHeader("Stick Calibration")) {
        ImGui::Indent(10);
        ImGui::TextDisabled("Connect a JoyCon first, then calibrate below.");
        ImGui::Spacing();

        bool hasAny = !g_singlePlayers.empty() || !g_dualPlayers.empty() || !g_proPlayers.empty();
        if (!hasAny) {
            ImGui::TextColored({1.f,0.6f,0.2f,1.f},"No controllers connected.");
        } else {
            ImGui::Text("Stick to calibrate:");
            ImGui::SameLine();
            if (ImGui::Button("Left Stick##cal")) ResetCalib(true);
            ImGui::SameLine();
            if (ImGui::Button("Right Stick##cal")) ResetCalib(false);
        }

        const auto& cal=GetActiveCalibration();
        ImGui::Spacing();
        ImGui::Text("Left:  center(%d,%d)  X[%d-%d]  Y[%d-%d]",
            cal.leftStick.centerX,cal.leftStick.centerY,
            cal.leftStick.minX,cal.leftStick.maxX,cal.leftStick.minY,cal.leftStick.maxY);
        ImGui::Text("Right: center(%d,%d)  X[%d-%d]  Y[%d-%d]",
            cal.rightStick.centerX,cal.rightStick.centerY,
            cal.rightStick.minX,cal.rightStick.maxX,cal.rightStick.minY,cal.rightStick.maxY);
        ImGui::Unindent(10); ImGui::Spacing();
    }

    bool hasProActive = !g_proPlayers.empty();
    if (hasProActive && ImGui::CollapsingHeader("GL/GR Layouts")) {
        ImGui::Indent(10);
        ImGui::Text("Active: %s  (GL: %s, GR: %s)",
            g_proConfig.layouts[g_proConfig.activeLayoutIndex].name,
            BtnMapStr(g_proConfig.layouts[g_proConfig.activeLayoutIndex].glMapping),
            BtnMapStr(g_proConfig.layouts[g_proConfig.activeLayoutIndex].grMapping));
        ImGui::SameLine();
        if (ImGui::Button("Manage Layouts")) g_showLayoutManager = true;
        ImGui::SameLine();
        if (ImGui::Button("Next Layout")) {
            g_proConfig.activeLayoutIndex = (g_proConfig.activeLayoutIndex+1) % (int)g_proConfig.layouts.size();
            SaveProConfig();
        }
        ImGui::Unindent(10); ImGui::Spacing();
    }

    if (ImGui::CollapsingHeader("Settings")) {
        ImGui::Indent(10);
        const char* policies[]={"Low Latency","Balanced 120Hz","Legacy 60Hz"};
        int pol=(int)g_opts.updatePolicy;
        ImGui::SetNextItemWidth(200);
        if (ImGui::Combo("Update Policy##run",&pol,policies,3))
            g_opts.updatePolicy=(UpdatePolicy)pol;
        ImGui::Unindent(10); ImGui::Spacing();
    }

    if (ImGui::CollapsingHeader("Log", ImGuiTreeNodeFlags_DefaultOpen))
    {
        static std::string logBuffer;
        static size_t lastLogCount = 0;
    
        if (lastLogCount != g_logLines.size())
        {
            std::lock_guard<std::mutex> lk(g_logMutex);
    
            logBuffer.clear();
    
            for (auto& line : g_logLines)
            {
                logBuffer += line;
                logBuffer += "\r\n";
            }
    
            lastLogCount = g_logLines.size();
        }
    
        if (ImGui::Button("Copy Log"))
        {
            ImGui::SetClipboardText(logBuffer.c_str());
        }
    
        ImGui::InputTextMultiline(
            "##log",
            logBuffer.data(),
            logBuffer.capacity() + 1,
            ImVec2(-FLT_MIN, 150),
            ImGuiInputTextFlags_ReadOnly
        );
    }

    if (g_openLayoutManager.load()) { g_openLayoutManager.store(false); g_showLayoutManager=true; }

    DrawCalibWizard();
    DrawLayoutManager();

    ImGui::End();
}

int WINAPI WinMain(HINSTANCE hInst, HINSTANCE, LPSTR, int) {
    init_apartment();
    LoadProConfig();
    LoadCalibrationProfiles("calibration.json");

    if (g_playerConfigs.empty()) g_playerConfigs.push_back({});

    WNDCLASSEXW wc{sizeof(WNDCLASSEXW), CS_CLASSDC, WndProc, 0L, 0L, hInst,
                   nullptr, nullptr, nullptr, nullptr, L"joycon2cpp", nullptr};
    RegisterClassExW(&wc);
    g_hwnd = CreateWindowW(L"joycon2cpp", L"joycon2cpp",
                           WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT,
                           900, 620, nullptr, nullptr, hInst, nullptr);

    if (!CreateDX11(g_hwnd)) { DestroyWindow(g_hwnd); UnregisterClassW(wc.lpszClassName, hInst); return 1; }
    ShowWindow(g_hwnd, SW_SHOWDEFAULT);
    UpdateWindow(g_hwnd);

    IMGUI_CHECKVERSION();
    ImGui::CreateContext();
    ImGuiIO& io = ImGui::GetIO();
    io.ConfigFlags |= ImGuiConfigFlags_NavEnableKeyboard;

    ImGui::StyleColorsDark();
    auto& style = ImGui::GetStyle();
    style.WindowRounding = 6.f;
    style.FrameRounding  = 4.f;
    style.GrabRounding   = 4.f;
    style.WindowBorderSize = 0.f;
    style.Colors[ImGuiCol_WindowBg] = {0.08f,0.08f,0.10f,1.f};
    style.Colors[ImGuiCol_Header]   = {0.18f,0.35f,0.55f,1.f};

    ImGui_ImplWin32_Init(g_hwnd);
    ImGui_ImplDX11_Init(g_pd3dDevice, g_pd3dDeviceContext);

    bool done = false;
    while (!done) {
        MSG msg;
        while (PeekMessage(&msg, nullptr, 0U, 0U, PM_REMOVE)) {
            TranslateMessage(&msg); DispatchMessage(&msg);
            if (msg.message == WM_QUIT) done = true;
        }
        if (done) break;

        ImGui_ImplDX11_NewFrame();
        ImGui_ImplWin32_NewFrame();
        ImGui::NewFrame();

        switch (g_screen) {
            case AppScreen::Setup:      DrawSetupScreen();      break;
            case AppScreen::Connecting: DrawConnectingScreen(); break;
            case AppScreen::Running:    DrawRunningScreen();    break;
        }

        ImGui::Render();
        const float cc[4] = {0.06f,0.06f,0.08f,1.f};
        g_pd3dDeviceContext->OMSetRenderTargets(1, &g_mainRTV, nullptr);
        g_pd3dDeviceContext->ClearRenderTargetView(g_mainRTV, cc);
        ImGui_ImplDX11_RenderDrawData(ImGui::GetDrawData());
        g_pSwapChain->Present(1, 0);
    }

    g_shuttingDown.store(true);

    for (auto* p : g_singleRumbleCtxs) if (p) p->running.store(false);
    for (auto* p : g_dualRumbleCtxs)   if (p) p->running.store(false);
    for (auto* p : g_proRumbleCtxs)    if (p) p->running.store(false);

    for (auto* p : g_singleRumbleCtxs) { if (p && p->thread.joinable()) p->thread.join(); }
    for (auto* p : g_dualRumbleCtxs)   { if (p && p->thread.joinable()) p->thread.join(); }
    for (auto* p : g_proRumbleCtxs)    { if (p && p->thread.joinable()) p->thread.join(); }

    for (auto* p : g_singleRumbleCtxs) delete p;
    for (auto* p : g_dualRumbleCtxs)   delete p;
    for (auto* p : g_proRumbleCtxs)    delete p;
    g_singleRumbleCtxs.clear();
    g_dualRumbleCtxs.clear();
    g_proRumbleCtxs.clear();

    for (auto& dp : g_dualPlayers) {
        if (!dp) continue;
        dp->running.store(false);
        if (dp->sharedState) dp->sharedState->cv.notify_all();
    }
    for (auto& dp : g_dualPlayers) {
        if (dp && dp->updateThread.joinable()) dp->updateThread.join();
    }

    if (g_vigem) {
        for (auto& dp : g_dualPlayers) {
            if (dp && dp->ds4Controller) {
                vigem_target_remove(g_vigem, dp->ds4Controller);
                vigem_target_free(dp->ds4Controller);
                dp->ds4Controller = nullptr;
            }
        }
        for (auto& sp : g_singlePlayers) {
            if (sp.ds4Controller) {
                vigem_target_remove(g_vigem, sp.ds4Controller);
                vigem_target_free(sp.ds4Controller);
                sp.ds4Controller = nullptr;
            }
        }
        for (auto& pp : g_proPlayers) {
            if (pp.ds4Controller) {
                vigem_target_remove(g_vigem, pp.ds4Controller);
                vigem_target_free(pp.ds4Controller);
                pp.ds4Controller = nullptr;
            }
        }

        vigem_disconnect(g_vigem);
        vigem_free(g_vigem);
        g_vigem = nullptr;
    }

    g_dsuServer.Stop();

    ImGui_ImplDX11_Shutdown();
    ImGui_ImplWin32_Shutdown();
    ImGui::DestroyContext();
    CleanupDX11();
    if (g_hwnd) {
        DestroyWindow(g_hwnd);
        g_hwnd = nullptr;
    }
    UnregisterClassW(wc.lpszClassName, hInst);
    uninit_apartment();
    return 0;
}
