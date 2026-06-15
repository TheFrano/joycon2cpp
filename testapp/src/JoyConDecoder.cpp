#include "JoyConDecoder.h"
#include <cmath>
#include <algorithm>
#include <ViGEm/Client.h>
#include <ViGEm/Common.h>
#include <cstdint>
#include <cstdio>
#include <vector>
#include <fstream>
#include <iostream>

int16_t to_signed_16(uint8_t lsb, uint8_t msb) {
    return static_cast<int16_t>((msb << 8) | lsb);
}

static std::vector<CalibrationProfile> g_calibrationProfiles;
static int g_activeCalibrationIndex = 0;

static CalibrationProfile MakeDefaultProfile()
{
    CalibrationProfile p;
    p.name = "Default";
    return p;
}

void ExtractRawStick(const std::vector<uint8_t>& buffer, bool isLeft, int& outX, int& outY)
{
    if (buffer.size() < 16) {
        outX = 2048;
        outY = 2048;
        return;
    }
    const uint8_t* data = isLeft ? &buffer[10] : &buffer[13];
    outX = ((data[1] & 0x0F) << 8) | data[0];
    outY = (data[2] << 4) | ((data[1] & 0xF0) >> 4);
}

void LoadCalibrationProfiles(const std::string& path)
{
    std::ifstream file(path);
    if (!file.is_open()) {
        if (g_calibrationProfiles.empty()) {
            g_calibrationProfiles.push_back(MakeDefaultProfile());
            g_activeCalibrationIndex = 0;
        }
        return;
    }

    std::string content((std::istreambuf_iterator<char>(file)), std::istreambuf_iterator<char>());
    file.close();

    g_calibrationProfiles.clear();
    g_activeCalibrationIndex = 0;

    size_t activePos = content.find("\"activeIndex\"");
    if (activePos != std::string::npos) {
        size_t colon = content.find(':', activePos);
        size_t end = content.find_first_of(",\n}", colon);
        try { g_activeCalibrationIndex = std::stoi(content.substr(colon + 1, end - colon - 1)); } catch (...) {}
    }

    size_t profilesPos = content.find("\"profiles\"");
    if (profilesPos == std::string::npos) {
        g_calibrationProfiles.push_back(MakeDefaultProfile());
        g_activeCalibrationIndex = 0;
        return;
    }

    size_t arrayStart = content.find('[', profilesPos);
    if (arrayStart == std::string::npos) {
        g_calibrationProfiles.push_back(MakeDefaultProfile());
        g_activeCalibrationIndex = 0;
        return;
    }

    size_t pos = arrayStart + 1;
    while (pos < content.size()) {
        size_t objStart = content.find('{', pos);
        if (objStart == std::string::npos) break;

        int depth = 1;
        size_t objEnd = objStart + 1;
        while (objEnd < content.size() && depth > 0) {
            if (content[objEnd] == '{') ++depth;
            else if (content[objEnd] == '}') --depth;
            ++objEnd;
        }

        std::string obj = content.substr(objStart, objEnd - objStart);
        CalibrationProfile p;

        size_t namePos = obj.find("\"name\"");
        if (namePos != std::string::npos) {
            size_t q1 = obj.find('"', namePos + 6);
            size_t q2 = obj.find('"', q1 + 1);
            if (q1 != std::string::npos && q2 != std::string::npos)
                p.name = obj.substr(q1 + 1, q2 - q1 - 1);
        }

        auto parseStick = [&](const std::string& sectionKey, StickCalibration& cal) {
            size_t spos = obj.find("\"" + sectionKey + "\"");
            if (spos == std::string::npos) return;
            size_t bstart = obj.find('{', spos);
            size_t bend = obj.find('}', bstart);
            if (bstart == std::string::npos || bend == std::string::npos) return;
            std::string section = obj.substr(bstart, bend - bstart + 1);

            auto readInt = [&](const std::string& key, int def) -> int {
                size_t kpos = section.find("\"" + key + "\"");
                if (kpos == std::string::npos) return def;
                size_t cpos = section.find(':', kpos);
                if (cpos == std::string::npos) return def;
                size_t epos = section.find_first_of(",}", cpos);
                try { return std::stoi(section.substr(cpos + 1, epos - cpos - 1)); } catch (...) { return def; }
            };

            cal.centerX = readInt("centerX", 2048);
            cal.centerY = readInt("centerY", 2048);
            cal.minX    = readInt("minX", 0);
            cal.maxX    = readInt("maxX", 4095);
            cal.minY    = readInt("minY", 0);
            cal.maxY    = readInt("maxY", 4095);
        };

        parseStick("leftStick", p.leftStick);
        parseStick("rightStick", p.rightStick);
        g_calibrationProfiles.push_back(p);
        pos = objEnd;
    }

    if (g_calibrationProfiles.empty()) {
        g_calibrationProfiles.push_back(MakeDefaultProfile());
        g_activeCalibrationIndex = 0;
    }
    if (g_activeCalibrationIndex < 0 || g_activeCalibrationIndex >= static_cast<int>(g_calibrationProfiles.size()))
        g_activeCalibrationIndex = 0;
}

void SaveCalibrationProfiles(const std::string& path)
{
    std::ofstream file(path);
    if (!file.is_open()) {
        std::cerr << "Failed to save calibration profiles to " << path << "\n";
        return;
    }

    auto writeStick = [&](const std::string& key, const StickCalibration& cal) {
        file << "      \"" << key << "\": {\n";
        file << "        \"centerX\": " << cal.centerX << ",\n";
        file << "        \"centerY\": " << cal.centerY << ",\n";
        file << "        \"minX\": "    << cal.minX    << ",\n";
        file << "        \"maxX\": "    << cal.maxX    << ",\n";
        file << "        \"minY\": "    << cal.minY    << ",\n";
        file << "        \"maxY\": "    << cal.maxY    << "\n";
        file << "      }";
    };

    file << "{\n";
    file << "  \"activeIndex\": " << g_activeCalibrationIndex << ",\n";
    file << "  \"profiles\": [\n";
    for (size_t i = 0; i < g_calibrationProfiles.size(); ++i) {
        const auto& p = g_calibrationProfiles[i];
        file << "    {\n";
        file << "      \"name\": \"" << p.name << "\",\n";
        writeStick("leftStick", p.leftStick);
        file << ",\n";
        writeStick("rightStick", p.rightStick);
        file << "\n";
        file << "    }";
        if (i + 1 < g_calibrationProfiles.size()) file << ",";
        file << "\n";
    }
    file << "  ]\n}\n";
    file.close();
    std::cout << "Calibration profiles saved to " << path << "\n";
}

const std::vector<CalibrationProfile>& GetCalibrationProfiles()
{
    return g_calibrationProfiles;
}

void AddCalibrationProfile(const CalibrationProfile& profile)
{
    g_calibrationProfiles.push_back(profile);
}

void DeleteCalibrationProfile(int index)
{
    if (index < 0 || index >= static_cast<int>(g_calibrationProfiles.size())) return;
    if (g_calibrationProfiles.size() <= 1) return;
    g_calibrationProfiles.erase(g_calibrationProfiles.begin() + index);
    if (g_activeCalibrationIndex >= static_cast<int>(g_calibrationProfiles.size()))
        g_activeCalibrationIndex = static_cast<int>(g_calibrationProfiles.size()) - 1;
}

int GetActiveCalibrationIndex()
{
    return g_activeCalibrationIndex;
}

void SetActiveCalibrationIndex(int index)
{
    if (index >= 0 && index < static_cast<int>(g_calibrationProfiles.size()))
        g_activeCalibrationIndex = index;
}

const CalibrationProfile& GetActiveCalibration()
{
    if (g_calibrationProfiles.empty()) {
        static CalibrationProfile def = MakeDefaultProfile();
        return def;
    }
    int idx = g_activeCalibrationIndex;
    if (idx < 0 || idx >= static_cast<int>(g_calibrationProfiles.size())) idx = 0;
    return g_calibrationProfiles[idx];
}

namespace {
void ApplyMotionToReport(DS4_REPORT_EX& report, const MotionData& motion)
{
    report.Report.wAccelX = motion.accelX;
    report.Report.wAccelY = motion.accelY;
    report.Report.wAccelZ = motion.accelZ;
    report.Report.wGyroX  = motion.gyroX;
    report.Report.wGyroY  = motion.gyroY;
    report.Report.wGyroZ  = motion.gyroZ;
}

float ApplyCalibratedAxis(int raw, int center, int minVal, int maxVal)
{
    if (minVal >= maxVal) {
        minVal = 300;
        maxVal = 3800;
        center = std::clamp(center, minVal, maxVal);
    }

    float result;
    if (raw >= center) {
        int range = maxVal - center;
        result = (range > 0) ? static_cast<float>(raw - center) / static_cast<float>(range) : 0.0f;
    } else {
        int range = center - minVal;
        result = (range > 0) ? -static_cast<float>(center - raw) / static_cast<float>(range) : 0.0f;
    }
    return std::clamp(result, -1.0f, 1.0f);
}
}

constexpr uint32_t BUTTON_A_MASK_RIGHT    = 0x000800;
constexpr uint32_t BUTTON_B_MASK_RIGHT    = 0x000200;
constexpr uint32_t BUTTON_X_MASK_RIGHT    = 0x000400;
constexpr uint32_t BUTTON_Y_MASK_RIGHT    = 0x000100;
constexpr uint32_t BUTTON_PLUS_MASK_RIGHT = 0x000002;
constexpr uint32_t BUTTON_R_MASK_RIGHT    = 0x004000;
constexpr uint32_t BUTTON_STICK_MASK_RIGHT = 0x000004;

constexpr uint32_t BUTTON_UP_MASK_LEFT    = 0x000002;
constexpr uint32_t BUTTON_DOWN_MASK_LEFT  = 0x000001;
constexpr uint32_t BUTTON_LEFT_MASK_LEFT  = 0x000008;
constexpr uint32_t BUTTON_RIGHT_MASK_LEFT = 0x000004;
constexpr uint32_t BUTTON_MINUS_MASK_LEFT = 0x000100;
constexpr uint32_t BUTTON_L_MASK_LEFT     = 0x000040;
constexpr uint32_t BUTTON_STICK_MASK_LEFT = 0x000800;

StickData DecodeJoystick(const std::vector<uint8_t>& buffer, JoyConSide side, JoyConOrientation orientation)
{
    if (buffer.size() < 16) {
        return { 0, 0, 0, 0 };
    }

    bool isLeft = (side == JoyConSide::Left);
    bool upright = (orientation == JoyConOrientation::Upright);

    int x_raw, y_raw;
    ExtractRawStick(buffer, isLeft, x_raw, y_raw);

    const StickCalibration& cal = isLeft
        ? GetActiveCalibration().leftStick
        : GetActiveCalibration().rightStick;

    float x = ApplyCalibratedAxis(x_raw, cal.centerX, cal.minX, cal.maxX);
    float y = ApplyCalibratedAxis(y_raw, cal.centerY, cal.minY, cal.maxY);

    if (!upright) {
        float tx = x, ty = y;
        x = isLeft ? -ty : ty;
        y = isLeft ? tx : -tx;
    }

    const float deadzone = 0.08f;
    if (std::abs(x) < deadzone && std::abs(y) < deadzone) {
        return { 0, 0, 0, 0 };
    }

    x = std::clamp(x * 1.7f, -1.0f, 1.0f);
    y = std::clamp(y * 1.7f, -1.0f, 1.0f);

    int16_t outX = static_cast<int16_t>(x * 32767);
    int16_t outY = static_cast<int16_t>(-y * 32767);

    return { outX, outY, 0, 0 };
}

static std::pair<int16_t, int16_t> decode_joystick(const std::vector<uint8_t>& buffer, bool isLeft, bool upright)
{
    auto res = DecodeJoystick(buffer, isLeft ? JoyConSide::Left : JoyConSide::Right, upright ? JoyConOrientation::Upright : JoyConOrientation::Sideways);
    return { res.x, res.y };
}

std::pair<uint16_t, uint16_t> DecodeMouseCoords(const std::vector<uint8_t>& buffer)
{
    if (buffer.size() < 0x18) return { 960, 471 };

    int16_t raw_x = to_signed_16(buffer[0x10], buffer[0x11]);
    int16_t raw_y = to_signed_16(buffer[0x12], buffer[0x13]);

    float norm_x = std::clamp(raw_x / 32767.0f, -1.0f, 1.0f);
    float norm_y = std::clamp(raw_y / 32767.0f, -1.0f, 1.0f);

    uint16_t x = static_cast<uint16_t>((norm_x + 1.0f) * 0.5f * 1920);
    uint16_t y = static_cast<uint16_t>((1.0f - (norm_y + 1.0f) * 0.5f) * 943);

    return { x, y };
}

void EncodeDS4Touch(DS4_TOUCH& touch, uint8_t trackingId, uint16_t x, uint16_t y)
{
    touch.bIsUpTrackingNum1 = trackingId & 0x7F;
    touch.bTouchData1[0] = x & 0xFF;
    touch.bTouchData1[1] = ((x >> 8) & 0x0F) | ((y & 0x0F) << 4);
    touch.bTouchData1[2] = (y >> 4) & 0xFF;
}

static void decode_triggers_shoulders(uint32_t state, bool isLeft, bool upright,
    BYTE& leftTrigger, BYTE& rightTrigger,
    bool& leftShoulder, bool& rightShoulder)
{
    leftTrigger  = (state & 0x000080) ? 255 : 0;
    rightTrigger = (state & 0x008000) ? 255 : 0;

    if (upright) {
        leftShoulder  = (state & 0x000040) != 0;
        rightShoulder = (state & 0x004000) != 0;
    } else {
        leftShoulder  = (state & (isLeft ? 0x000020 : 0x002000)) != 0;
        rightShoulder = (state & (isLeft ? 0x000010 : 0x001000)) != 0;
    }
}

MotionData DecodeMotionRaw(const std::vector<uint8_t>& buffer)
{
    MotionData motion{};
    if (buffer.size() < 0x3C) return motion;

    motion.accelX = to_signed_16(buffer[0x30], buffer[0x31]);
    motion.accelY = to_signed_16(buffer[0x32], buffer[0x33]);
    motion.accelZ = to_signed_16(buffer[0x34], buffer[0x35]);
    motion.gyroX  = to_signed_16(buffer[0x36], buffer[0x37]);
    motion.gyroY  = to_signed_16(buffer[0x38], buffer[0x39]);
    motion.gyroZ  = to_signed_16(buffer[0x3A], buffer[0x3B]);

    return motion;
}

MotionData DecodeMotion(const std::vector<uint8_t>& buffer)
{
    return DecodeMotionRaw(buffer);
}

static std::pair<int16_t, int16_t> decode_calibrated_stick(const uint8_t* data, const StickCalibration& cal)
{
    if (!data) return { 0, 0 };

    int x_raw = ((data[1] & 0x0F) << 8) | data[0];
    int y_raw = (data[2] << 4) | ((data[1] & 0xF0) >> 4);

    float x = ApplyCalibratedAxis(x_raw, cal.centerX, cal.minX, cal.maxX);
    float y = ApplyCalibratedAxis(y_raw, cal.centerY, cal.minY, cal.maxY);

    constexpr float deadzone = 0.08f;
    if (std::abs(x) < deadzone && std::abs(y) < deadzone) return { 0, 0 };

    x = std::clamp(x * 1.7f, -1.0f, 1.0f);
    y = std::clamp(y * 1.7f, -1.0f, 1.0f);

    return {
        static_cast<int16_t>(x * 32767),
        static_cast<int16_t>(y * 32767)
    };
}

DS4_REPORT_EX GenerateDS4Report(const std::vector<uint8_t>& buffer, JoyConSide side, JoyConOrientation orientation)
{
    DS4_REPORT_EX report{};
    DS4_REPORT_INIT(reinterpret_cast<PDS4_REPORT>(&report.Report));

    if (buffer.size() < 0x3C) return report;

    bool isLeft = (side == JoyConSide::Left);
    bool upright = (orientation == JoyConOrientation::Upright);

    int btnOffset = isLeft ? 4 : 3;
    uint32_t state = (buffer[btnOffset] << 16) | (buffer[btnOffset + 1] << 8) | buffer[btnOffset + 2];

    auto [stickX, stickY] = decode_joystick(buffer, isLeft, upright);

    if (isLeft) {
        bool up    = (state & BUTTON_UP_MASK_LEFT) != 0;
        bool down  = (state & BUTTON_DOWN_MASK_LEFT) != 0;
        bool left  = (state & BUTTON_LEFT_MASK_LEFT) != 0;
        bool right = (state & BUTTON_RIGHT_MASK_LEFT) != 0;

        uint8_t dpad = DS4_BUTTON_DPAD_NONE;
        if      (up && left)   dpad = DS4_BUTTON_DPAD_NORTHWEST;
        else if (up && right)  dpad = DS4_BUTTON_DPAD_NORTHEAST;
        else if (down && left) dpad = DS4_BUTTON_DPAD_SOUTHWEST;
        else if (down && right)dpad = DS4_BUTTON_DPAD_SOUTHEAST;
        else if (up)           dpad = DS4_BUTTON_DPAD_NORTH;
        else if (down)         dpad = DS4_BUTTON_DPAD_SOUTH;
        else if (left)         dpad = DS4_BUTTON_DPAD_WEST;
        else if (right)        dpad = DS4_BUTTON_DPAD_EAST;

        DS4_SET_DPAD(reinterpret_cast<PDS4_REPORT>(&report.Report), static_cast<DS4_DPAD_DIRECTIONS>(dpad));

        if (state & BUTTON_MINUS_MASK_LEFT)  report.Report.wButtons |= DS4_BUTTON_SHARE;
        if (state & BUTTON_L_MASK_LEFT)      report.Report.wButtons |= DS4_BUTTON_SHOULDER_LEFT;
        if (state & BUTTON_STICK_MASK_LEFT)  report.Report.wButtons |= DS4_BUTTON_THUMB_LEFT;
    } else {
        DS4_SET_DPAD(reinterpret_cast<PDS4_REPORT>(&report.Report), DS4_BUTTON_DPAD_NONE);

        if (state & BUTTON_A_MASK_RIGHT)     report.Report.wButtons |= DS4_BUTTON_CIRCLE;
        if (state & BUTTON_B_MASK_RIGHT)     report.Report.wButtons |= DS4_BUTTON_TRIANGLE;
        if (state & BUTTON_X_MASK_RIGHT)     report.Report.wButtons |= DS4_BUTTON_CROSS;
        if (state & BUTTON_Y_MASK_RIGHT)     report.Report.wButtons |= DS4_BUTTON_SQUARE;
        if (state & BUTTON_PLUS_MASK_RIGHT)  report.Report.wButtons |= DS4_BUTTON_OPTIONS;
        if (state & BUTTON_R_MASK_RIGHT)     report.Report.wButtons |= DS4_BUTTON_SHOULDER_RIGHT;
        if (state & BUTTON_STICK_MASK_RIGHT) report.Report.wButtons |= DS4_BUTTON_THUMB_RIGHT;
    }

    auto [touchX, touchY] = DecodeMouseCoords(buffer);
    report.Report.bTouchPacketsN = 1;
    report.Report.sCurrentTouch.bPacketCounter++;
    EncodeDS4Touch(report.Report.sCurrentTouch, 1, touchX, touchY);

    BYTE leftTrigger = 0, rightTrigger = 0;
    bool leftShoulder = false, rightShoulder = false;
    decode_triggers_shoulders(state, isLeft, upright, leftTrigger, rightTrigger, leftShoulder, rightShoulder);

    if (leftShoulder)  report.Report.wButtons |= DS4_BUTTON_SHOULDER_LEFT;
    if (rightShoulder) report.Report.wButtons |= DS4_BUTTON_SHOULDER_RIGHT;
    if (leftTrigger)   report.Report.wButtons |= DS4_BUTTON_TRIGGER_LEFT;
    if (rightTrigger)  report.Report.wButtons |= DS4_BUTTON_TRIGGER_RIGHT;

    report.Report.bThumbLX = static_cast<BYTE>((stickX / 32767.0f) * 127 + 128);
    report.Report.bThumbLY = static_cast<BYTE>((stickY / 32767.0f) * 127 + 128);

    ApplyMotionToReport(report, DecodeMotionRaw(buffer));

    return report;
}

DS4_REPORT_EX GenerateDualJoyConDS4Report(const std::vector<uint8_t>& leftBuffer, const std::vector<uint8_t>& rightBuffer, GyroSource gyroSource)
{
    DS4_REPORT_EX report{};
    DS4_REPORT_INIT(reinterpret_cast<PDS4_REPORT>(&report.Report));

    if (leftBuffer.size() < 0x3C && rightBuffer.size() < 0x3C) return report;

    DS4_REPORT_EX leftReport{};
    if (leftBuffer.size() >= 0x3C)
        leftReport = GenerateDS4Report(leftBuffer, JoyConSide::Left, JoyConOrientation::Upright);

    DS4_REPORT_EX rightReport{};
    if (rightBuffer.size() >= 0x3C)
        rightReport = GenerateDS4Report(rightBuffer, JoyConSide::Right, JoyConOrientation::Upright);

    USHORT leftDpad           = leftReport.Report.wButtons & 0xF;
    USHORT leftButtonsNoDpad  = leftReport.Report.wButtons & ~0xF;
    USHORT rightButtonsNoDpad = rightReport.Report.wButtons & ~0xF;

    report.Report.wButtons = (leftButtonsNoDpad | rightButtonsNoDpad) | leftDpad;
    report.Report.bSpecial = leftReport.Report.bSpecial | rightReport.Report.bSpecial;

    auto [x1, y1] = DecodeMouseCoords(leftBuffer);
    auto [x2, y2] = DecodeMouseCoords(rightBuffer);

    report.Report.bTouchPacketsN = 1;
    report.Report.sCurrentTouch.bPacketCounter++;
    EncodeDS4Touch(report.Report.sCurrentTouch, 1, x1, y1);
    report.Report.sCurrentTouch.bIsUpTrackingNum2 = 2;
    report.Report.sCurrentTouch.bTouchData2[0] = x2 & 0xFF;
    report.Report.sCurrentTouch.bTouchData2[1] = ((x2 >> 8) & 0x0F) | ((y2 & 0x0F) << 4);
    report.Report.sCurrentTouch.bTouchData2[2] = (y2 >> 4) & 0xFF;

    uint32_t leftState  = (leftBuffer[4]  << 16) | (leftBuffer[5]  << 8) | leftBuffer[6];
    uint32_t rightState = (rightBuffer[3] << 16) | (rightBuffer[4] << 8) | rightBuffer[5];

    BYTE lt = 0, rt = 0;
    bool ls = false, rs = false;

    decode_triggers_shoulders(leftState, true, true, lt, rt, ls, rs);
    report.Report.bTriggerL = lt;
    if (ls) report.Report.wButtons |= DS4_BUTTON_SHOULDER_LEFT;
    if (lt) report.Report.wButtons |= DS4_BUTTON_TRIGGER_LEFT;

    decode_triggers_shoulders(rightState, false, true, lt, rt, ls, rs);
    report.Report.bTriggerR = rt;
    if (rs) report.Report.wButtons |= DS4_BUTTON_SHOULDER_RIGHT;
    if (rt) report.Report.wButtons |= DS4_BUTTON_TRIGGER_RIGHT;

    report.Report.bThumbLX = leftReport.Report.bThumbLX;
    report.Report.bThumbLY = leftReport.Report.bThumbLY;
    report.Report.bThumbRX = rightReport.Report.bThumbLX;
    report.Report.bThumbRY = rightReport.Report.bThumbLY;

    switch (gyroSource) {
        case GyroSource::Left:
            report.Report.wAccelX = leftReport.Report.wAccelX;
            report.Report.wAccelY = leftReport.Report.wAccelY;
            report.Report.wAccelZ = leftReport.Report.wAccelZ;
            report.Report.wGyroX  = leftReport.Report.wGyroX;
            report.Report.wGyroY  = leftReport.Report.wGyroY;
            report.Report.wGyroZ  = leftReport.Report.wGyroZ;
            break;
        case GyroSource::Right:
            report.Report.wAccelX = rightReport.Report.wAccelX;
            report.Report.wAccelY = rightReport.Report.wAccelY;
            report.Report.wAccelZ = rightReport.Report.wAccelZ;
            report.Report.wGyroX  = rightReport.Report.wGyroX;
            report.Report.wGyroY  = rightReport.Report.wGyroY;
            report.Report.wGyroZ  = rightReport.Report.wGyroZ;
            break;
        case GyroSource::Both:
        default: {
            auto combine_16 = [](int16_t a, int16_t b) -> int16_t {
                if (a == 0) return b;
                if (b == 0) return a;
                return static_cast<int16_t>((a / 2) + (b / 2));
            };
            report.Report.wAccelX = combine_16(leftReport.Report.wAccelX, rightReport.Report.wAccelX);
            report.Report.wAccelY = combine_16(leftReport.Report.wAccelY, rightReport.Report.wAccelY);
            report.Report.wAccelZ = combine_16(leftReport.Report.wAccelZ, rightReport.Report.wAccelZ);
            report.Report.wGyroX  = combine_16(leftReport.Report.wGyroX,  rightReport.Report.wGyroX);
            report.Report.wGyroY  = combine_16(leftReport.Report.wGyroY,  rightReport.Report.wGyroY);
            report.Report.wGyroZ  = combine_16(leftReport.Report.wGyroZ,  rightReport.Report.wGyroZ);
            break;
        }
    }

    return report;
}

constexpr uint64_t BUTTON_A_MASK      = 0x000800000000;
constexpr uint64_t BUTTON_B_MASK      = 0x000400000000;
constexpr uint64_t BUTTON_X_MASK      = 0x000200000000;
constexpr uint64_t BUTTON_Y_MASK      = 0x000100000000;
constexpr uint64_t BUTTON_R_SHOULDER  = 0x004000000000;
constexpr uint64_t BUTTON_L_SHOULDER  = 0x000000400000;
constexpr uint64_t BUTTON_DPAD_UP     = 0x000000020000;
constexpr uint64_t BUTTON_DPAD_RIGHT  = 0x000000040000;
constexpr uint64_t BUTTON_DPAD_DOWN   = 0x000000010000;
constexpr uint64_t BUTTON_DPAD_LEFT   = 0x000000080000;
constexpr uint64_t BUTTON_GUIDE       = 0x000010000000;
constexpr uint64_t BUTTON_BACK        = 0x000001000000;
constexpr uint64_t BUTTON_START       = 0x000002000000;
constexpr uint64_t BUTTON_R_THUMB     = 0x000004000000;
constexpr uint64_t BUTTON_L_THUMB     = 0x000008000000;
constexpr uint64_t TRIGGER_LT_MASK    = 0x000000800000;
constexpr uint64_t TRIGGER_RT_MASK    = 0x008000000000;

DS4_REPORT_EX GenerateProControllerReport(const std::vector<uint8_t>& buffer)
{
    DS4_REPORT_EX report{};
    DS4_REPORT_INIT(reinterpret_cast<PDS4_REPORT>(&report.Report));

    if (buffer.size() < 0x3C) return report;

    uint64_t state = 0;
    for (int i = 3; i <= 8; ++i) state = (state << 8) | buffer[i];

    if (state & BUTTON_A_MASK)     report.Report.wButtons |= DS4_BUTTON_CIRCLE;
    if (state & BUTTON_B_MASK)     report.Report.wButtons |= DS4_BUTTON_CROSS;
    if (state & BUTTON_X_MASK)     report.Report.wButtons |= DS4_BUTTON_TRIANGLE;
    if (state & BUTTON_Y_MASK)     report.Report.wButtons |= DS4_BUTTON_SQUARE;
    if (state & BUTTON_L_SHOULDER) report.Report.wButtons |= DS4_BUTTON_SHOULDER_LEFT;
    if (state & BUTTON_R_SHOULDER) report.Report.wButtons |= DS4_BUTTON_SHOULDER_RIGHT;
    if (state & BUTTON_L_THUMB)    report.Report.wButtons |= DS4_BUTTON_THUMB_LEFT;
    if (state & BUTTON_R_THUMB)    report.Report.wButtons |= DS4_BUTTON_THUMB_RIGHT;
    if (state & BUTTON_BACK)       report.Report.bSpecial |= DS4_SPECIAL_BUTTON_TOUCHPAD;
    if (state & BUTTON_START)      report.Report.wButtons |= DS4_BUTTON_OPTIONS;
    if (state & BUTTON_GUIDE)      report.Report.bSpecial |= DS4_SPECIAL_BUTTON_PS;

    bool up    = (state & BUTTON_DPAD_UP)    != 0;
    bool down  = (state & BUTTON_DPAD_DOWN)  != 0;
    bool left  = (state & BUTTON_DPAD_LEFT)  != 0;
    bool right = (state & BUTTON_DPAD_RIGHT) != 0;

    uint8_t dpad = DS4_BUTTON_DPAD_NONE;
    if      (up && left)   dpad = DS4_BUTTON_DPAD_NORTHWEST;
    else if (up && right)  dpad = DS4_BUTTON_DPAD_NORTHEAST;
    else if (down && left) dpad = DS4_BUTTON_DPAD_SOUTHWEST;
    else if (down && right)dpad = DS4_BUTTON_DPAD_SOUTHEAST;
    else if (up)           dpad = DS4_BUTTON_DPAD_NORTH;
    else if (down)         dpad = DS4_BUTTON_DPAD_SOUTH;
    else if (left)         dpad = DS4_BUTTON_DPAD_WEST;
    else if (right)        dpad = DS4_BUTTON_DPAD_EAST;

    DS4_SET_DPAD(reinterpret_cast<PDS4_REPORT>(&report.Report), static_cast<DS4_DPAD_DIRECTIONS>(dpad));

    report.Report.bTriggerL = (state & TRIGGER_LT_MASK) ? 255 : 0;
    report.Report.bTriggerR = (state & TRIGGER_RT_MASK) ? 255 : 0;

    const auto& cal = GetActiveCalibration();
    auto [lx, ly] = decode_calibrated_stick(&buffer[10], cal.leftStick);
    ly = -ly;
    auto [rx, ry] = decode_calibrated_stick(&buffer[13], cal.rightStick);
    ry = -ry;

    report.Report.bThumbLX = static_cast<uint8_t>((lx / 32767.0f) * 127 + 128);
    report.Report.bThumbLY = static_cast<uint8_t>((ly / 32767.0f) * 127 + 128);
    report.Report.bThumbRX = static_cast<uint8_t>((rx / 32767.0f) * 127 + 128);
    report.Report.bThumbRY = static_cast<uint8_t>((ry / 32767.0f) * 127 + 128);

    ApplyMotionToReport(report, DecodeMotionRaw(buffer));

    return report;
}

DS4_REPORT_EX GenerateNSOGCReport(const std::vector<uint8_t>& buffer)
{
    DS4_REPORT_EX report{};
    DS4_REPORT_INIT(reinterpret_cast<PDS4_REPORT>(&report.Report));

    if (buffer.size() < 0x3C) return report;

    uint64_t state = 0;
    for (int i = 3; i <= 8; ++i) state = (state << 8) | buffer[i];

    if (state & BUTTON_A_MASK)     report.Report.wButtons |= DS4_BUTTON_CIRCLE;
    if (state & BUTTON_B_MASK)     report.Report.wButtons |= DS4_BUTTON_TRIANGLE;
    if (state & BUTTON_X_MASK)     report.Report.wButtons |= DS4_BUTTON_CROSS;
    if (state & BUTTON_Y_MASK)     report.Report.wButtons |= DS4_BUTTON_SQUARE;
    if (state & BUTTON_L_SHOULDER) report.Report.wButtons |= DS4_BUTTON_SHOULDER_LEFT;
    if (state & BUTTON_R_SHOULDER) report.Report.wButtons |= DS4_BUTTON_SHOULDER_RIGHT;
    if (state & TRIGGER_LT_MASK)   report.Report.wButtons |= DS4_BUTTON_TRIGGER_LEFT;
    if (state & TRIGGER_RT_MASK)   report.Report.wButtons |= DS4_BUTTON_TRIGGER_RIGHT;
    if (state & BUTTON_L_THUMB)    report.Report.wButtons |= DS4_BUTTON_THUMB_LEFT;
    if (state & BUTTON_R_THUMB)    report.Report.wButtons |= DS4_BUTTON_THUMB_RIGHT;
    if (state & BUTTON_BACK)       report.Report.wButtons |= DS4_BUTTON_SHARE;
    if (state & BUTTON_START)      report.Report.wButtons |= DS4_BUTTON_OPTIONS;
    if (state & BUTTON_GUIDE)      report.Report.bSpecial |= DS4_SPECIAL_BUTTON_PS;

    bool up    = (state & BUTTON_DPAD_UP)    != 0;
    bool down  = (state & BUTTON_DPAD_DOWN)  != 0;
    bool left  = (state & BUTTON_DPAD_LEFT)  != 0;
    bool right = (state & BUTTON_DPAD_RIGHT) != 0;

    uint8_t dpad = DS4_BUTTON_DPAD_NONE;
    if      (up && left)   dpad = DS4_BUTTON_DPAD_NORTHWEST;
    else if (up && right)  dpad = DS4_BUTTON_DPAD_NORTHEAST;
    else if (down && left) dpad = DS4_BUTTON_DPAD_SOUTHWEST;
    else if (down && right)dpad = DS4_BUTTON_DPAD_SOUTHEAST;
    else if (up)           dpad = DS4_BUTTON_DPAD_NORTH;
    else if (down)         dpad = DS4_BUTTON_DPAD_SOUTH;
    else if (left)         dpad = DS4_BUTTON_DPAD_WEST;
    else if (right)        dpad = DS4_BUTTON_DPAD_EAST;

    DS4_SET_DPAD(reinterpret_cast<PDS4_REPORT>(&report.Report), static_cast<DS4_DPAD_DIRECTIONS>(dpad));

    report.Report.bTriggerL = buffer[0x3c];
    report.Report.bTriggerR = buffer[0x3d];

    const auto& cal = GetActiveCalibration();
    auto [lx, ly] = decode_calibrated_stick(&buffer[10], cal.leftStick);
    ly = -ly;
    auto [rx, ry] = decode_calibrated_stick(&buffer[13], cal.rightStick);
    ry = -ry;

    report.Report.bThumbLX = static_cast<uint8_t>((lx / 32767.0f) * 127 + 128);
    report.Report.bThumbLY = static_cast<uint8_t>((ly / 32767.0f) * 127 + 128);
    report.Report.bThumbRX = static_cast<uint8_t>((rx / 32767.0f) * 127 + 128);
    report.Report.bThumbRY = static_cast<uint8_t>((ry / 32767.0f) * 127 + 128);

    ApplyMotionToReport(report, DecodeMotionRaw(buffer));

    return report;
}

uint32_t ExtractButtonState(const std::vector<uint8_t>& buffer)
{
    if (buffer.size() < 6) return 0;
    return (buffer[3] << 16) | (buffer[4] << 8) | buffer[5];
}

std::pair<int16_t, int16_t> GetRawOpticalMouse(const std::vector<uint8_t>& buffer)
{
    if (buffer.size() < 0x18) return { 0, 0 };
    int16_t raw_x = to_signed_16(buffer[0x10], buffer[0x11]);
    int16_t raw_y = to_signed_16(buffer[0x12], buffer[0x13]);
    return { raw_x, raw_y };
}
