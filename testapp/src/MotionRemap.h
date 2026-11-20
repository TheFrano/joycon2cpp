#pragma once

#include <cstdint>

struct MotionAxisMap {
    int  src[3]    = { 0, 1, 2 };            
    bool invert[3] = { false, false, false }; 

    constexpr bool operator==(const MotionAxisMap& o) const {
        for (int i = 0; i < 3; ++i)
            if (src[i] != o.src[i] || invert[i] != o.invert[i]) return false;
        return true;
    }
    constexpr bool operator!=(const MotionAxisMap& o) const { return !(*this == o); }
};

constexpr MotionAxisMap MotionAxisMapRaw() { return MotionAxisMap{}; }

constexpr MotionAxisMap MotionAxisMapSwitchToDS4()
{
    MotionAxisMap m;
    m.src[0] = 0;                       
    m.src[1] = 2;                       
    m.src[2] = 1; m.invert[2] = true;   
    return m;
}

inline int16_t MotionNegateClamped(int16_t v)
{
    return v == INT16_MIN ? INT16_MAX : static_cast<int16_t>(-v);
}

inline void RemapMotionAxes(const int16_t in[3], int16_t out[3], const MotionAxisMap& m)
{
    for (int i = 0; i < 3; ++i) {
        int s = m.src[i];
        if (s < 0 || s > 2) s = i;
        out[i] = m.invert[i] ? MotionNegateClamped(in[s]) : in[s];
    }
}

constexpr uint32_t PackMotionAxisMap(const MotionAxisMap& m)
{
    uint32_t v = 0;
    for (int i = 0; i < 3; ++i)
        v |= (static_cast<uint32_t>(m.src[i] & 3) | (m.invert[i] ? 4u : 0u)) << (i * 3);
    return v;
}

constexpr MotionAxisMap UnpackMotionAxisMap(uint32_t v)
{
    MotionAxisMap m{};
    for (int i = 0; i < 3; ++i) {
        const uint32_t f = (v >> (i * 3)) & 7u;
        m.src[i]    = (f & 3u) > 2u ? i : static_cast<int>(f & 3u);
        m.invert[i] = (f & 4u) != 0;
    }
    return m;
}
