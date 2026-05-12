# joycon2cpp

A C++ executable that makes Switch 2 controllers into working PC controllers.

---

## DISCLAIMER

This project is **Windows-only**, primarily because the `ViGEmBus Driver` (used for virtual controller output) is exclusive to Windows.  
You're free to make your own macOS/Linux fork if you want.
---

## DEPENDENCIES

- [ViGEmBus drivers](https://github.com/ViGEm/ViGEmBus/releases/latest)
- [Microsoft Visual C++ Redistributable 2015–2022 (x64)](https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist)

---

## How do I use it?
- Download the exe in the source code and open it, or optionally build the source code and open the created exe file
- Pick your amount of players
- Pick everyone's controller
- If using a singular joycon you'll be asked if its Left or Right
- If using dual joycons itll ask you to pair one joycon then the other
- When its all done, you'll have SDL controllers ready for every player to use.
- Click the chat button to toggle mouse functions，3 adjustable modes for the mouse pointer.

>  Note: Bit layouts differ slightly between left and right Joy-Cons, so correct side pairing is important.
> 
---

## Building from source

If you want to build the project yourself, follow these instructions (Windows + Visual Studio):

 Requirements  

Make sure the following are installed via Visual Studio Installer:

 Visual Studio 2022 or newer
 Workload: Desktop development with C++
 Component: Windows 10 or 11 SDK
 Component: MSVC v14.x

 (Optional but useful) C++/WinRT  

1. Open the x64 Native Tools Command Prompt for VS (whatever your version is)

2. Go into the project's root directory and make a build folder:

   ```sh
   cd (path to the main directory of the project here)
   mkdir build
   cd build

3. Generate Visual Studio project files with CMake:
    ```sh
    cmake .. -G "Visual Studio 17 2022" -A x64
4. Build the project in Release mode:
    ```sh
    cmake --build . --config Release
5. The compiled executable will be located in:
    ```sh
    build\Release\testapp.exe

# Joy-Con 2 BLE Notification Research
(this section is for documenting the Joy-Con 2's behavior and function. if you're just looking to use the program, this section isnt important.)





This document outlines some findings related to Joy-Con 2 BLE input behavior. If you're developing or reverse-engineering Joy-Con 2, Pro Controller 2, or other supported Nintendo controllers over BLE, this may be useful.

##  Behavior Quirks

A notable quirk of these controllers is that if you attempt to connect or pair them repeatedly in a short time span, they may stop responding or fail to connect entirely for several minutes. This appears to be a controller-level cooldown behavior rather than an OS/BLE stack issue.

**If your controller stops connecting:**  
Wait a few minutes before trying again. It should recover on its own.

##  BLE Notification (with IMU enabled, Left Joy-Con)

Here’s an example notification received from a Joy-Con 2 via BLE, with the IMU command sent. (Pro Controller 2 and GC Controller notifications follow similar layouts but may shift certain fields.)

08670000000000e0ff0ffff77f23287a0000000000000000000000000000005f0e007907000000000001ce7b52010500beffb501ee0ffeff04000200000000


### Field Breakdown (based on known Joy-Con 2 layout)
huge thanks to [@german77](https://github.com/german77) for providing me with the notification layout below!!

| Offset | Size | Value              | Comment                      |
|--------|------|--------------------|------------------------------|
| `0x00` | 0x4  | Packet ID          | Sequence or timestamp        |
| `0x04` | 0x4  | Buttons            | Button state bitmap          |
| `0x08` | 0x3  | Left Stick         | 12-bit X/Y packed             |
| `0x0B` | 0x3  | Right Stick        | 12-bit X/Y packed   |
| `0x0E` | 0x2  | Mouse X            |              |
| `0x10` | 0x2  | Mouse Y            |                 |
| `0x12` | 0x2  | Mouse Unk          | Possibly extra motion data    |
| `0x14` | 0x2  | Mouse Distance     | Distance to IR/motion surface |
| `0x16` | 0x2  | Magnetometer X     |                              |
| `0x18` | 0x2  | Magnetometer Y     |                              |
| `0x1A` | 0x2  | Magnetometer Z     |                              |
| `0x1C` | 0x2  | Battery Voltage    | 1000 = 1V                     |
| `0x1E` | 0x2  | Battery Current    | 100 = 1mA                     |
| `0x20` | 0xE  | Reserved           | Undocumented region           |
| `0x2E` | 0x2  | Temperature        | `25°C + raw / 127`           |
| `0x30` | 0x2  | Accel X            | 4096 = 1G                     |
| `0x32` | 0x2  | Accel Y            |                              |
| `0x34` | 0x2  | Accel Z            |                              |
| `0x36` | 0x2  | Gyro X             | 48000 = 360°/s                |
| `0x38` | 0x2  | Gyro Y             |                              |
| `0x3A` | 0x2  | Gyro Z             |                              |
| `0x3C` | 0x1  | Analog Trigger L   |                              |
| `0x3D` | 0x1  | Analog Trigger R   |                              |

---

### 🧪 Field Example Breakdown

| Offset | Size | Field           | Raw Value     | Interpreted                  |
|--------|------|------------------|----------------|------------------------------|
| `0x00` | 4    | Packet ID        | `08 67 00 00`  | `0x00006708` → `26376`       |
| `0x04` | 4    | Buttons          | `00 00 00 00`  | No buttons pressed           |
| `0x08` | 3    | Left Stick       | `e0 ff 0f`     | X = `0x0FF0` = `4080`, Y = `0x0FE0` = `4064` |
| `0x0B` | 3    | Right Stick      | `ff f7 7f`     | Garbage on Left Joy-Con      |
| `0x2E` | 2    | Temperature      | `5f 0e`        | `0x0E5F` = `3679` → ~54°C     |
| `0x30` | 2    | Accel X          | `00 79`        | `0x7900` = `30976`           |
| `0x32` | 2    | Accel Y          | `07 00`        | `0x0007` = `7`               |
| `0x34` | 2    | Accel Z          | `00 00`        | `0`                          |
| `0x36` | 2    | Gyro X           | `01 ce`        | `0xCE01` = `52737`           |
| `0x38` | 2    | Gyro Y           | `7b 52`        | `0x527B` = `21115`           |
| `0x3A` | 2    | Gyro Z           | `01 05`        | `0x0501` = `1281`            |

---

### 📘 Notes

- Left Joy-Con **does not use Right Stick**, so data at `0x0B–0x0D` is typically junk.
- **Stick values** use 12-bit X/Y packed across 3 bytes:
  - X = upper 12 bits of first 1.5 bytes
  - Y = lower 12 bits of next 1.5 bytes
- **Accel/Gyro** fields are signed 16-bit:
  - Accelerometer: `4096 = 1G`
  - Gyroscope: `48000 = 360°/s`
- **Temperature**:  
  `25°C + (raw / 127)`  
  → `25 + (3679 / 127) ≈ 54°C`
- **Battery voltage**:  
Reported as millivolts. `3000` = 3.0V. If `0x0000`, likely unavailable at that time.

---

## Latency diagnostics

The test app can write an internal latency CSV while you compare update policies:

```powershell
.\testapp.exe --latency-test --update-policy low --latency-csv low.csv
.\testapp.exe --latency-test --update-policy balanced --latency-csv balanced.csv
.\testapp.exe --latency-test --update-policy legacy --latency-csv legacy.csv
```

Policies:

- `low` / `LowLatency`: send a ViGEm update for each BLE notification.
- `balanced` / `Balanced120Hz`: limit output to about 120 Hz.
- `legacy` / `Legacy60Hz`: limit output to about 60 Hz for comparison.

The CSV columns are:

```text
mode,controller_type,event_index,ble_delta_ms,buffer_age_left_ms,buffer_age_right_ms,decode_to_vigem_us,total_pipeline_us
```

For a repeatable manual comparison, run each policy with the same controller, keep it still for 10 seconds, then press one button 30 times at a steady rhythm. Repeat for Single Joy-Con, Dual Joy-Con, and Pro Controller. For perceived end-to-end latency, record the physical controller and gamepad-tester.com or Steam Input at 240 fps, count frames between the visible press and on-screen response, and convert with `latency_ms = frames / fps * 1000`.
