# -*- coding: utf-8 -*-
"""
生成《失控 OUT OF CONTROL》科技实验室氛围 BGM（WAV，30 秒可循环）
纯 Python 合成，无需第三方库。
"""
import math, wave, array, random, time

SR = 44100          # 采样率
DUR = 30.0          # 时长（秒），循环播放
N = int(SR * DUR)
OUT = r"C:\Users\Yan\WorkBuddy\2026-08-12-17-54-20\dungeon-guardian\bgm_ambient.wav"

random.seed(7)
TWO_PI = 2 * math.pi

# A 小调琶音序列（8 音，旋律感更强）
arp_notes = [110.0, 130.81, 164.81, 220.0, 164.81, 130.81, 196.0, 164.81]

def env_exp(t, k):
    return math.exp(-t * k)

buf = array.array('h')

t0 = time.time()
for i in range(N):
    t = i / SR

    # ---- 1. 低频 Drone（科技实验室底噪感）----
    drone = (0.15 * math.sin(TWO_PI * 55.0 * t) +      # A1
             0.09 * math.sin(TWO_PI * 110.0 * t) +     # A2
             0.05 * math.sin(TWO_PI * 164.81 * t))     # E3 五度
    # 慢速呼吸 LFO
    lfo = 0.72 + 0.28 * math.sin(TWO_PI * 0.08 * t)
    # 后半段加入 D2 增厚
    if t >= 12.0:
        drone += 0.06 * math.sin(TWO_PI * 73.42 * t)
    drone *= lfo

    # ---- 2. 琶音（柔和电子音色，快起慢衰）----
    arp_idx = int(t / 0.4) % len(arp_notes)
    arp_f = arp_notes[arp_idx]
    arp_tt = t % 0.4
    arp = 0.07 * env_exp(arp_tt, 7.0) * math.sin(TWO_PI * arp_f * t)

    # ---- 3. 高频脉冲（科技感滴滴声，每 2 秒）----
    pulse_t = t % 2.0
    pulse = 0.0
    if pulse_t < 0.12:
        pulse = 0.05 * env_exp(pulse_t, 40.0) * math.sin(TWO_PI * 1320.0 * t)

    # （已移除白噪声底噪——听感沙沙，去掉了）

    s = drone + arp + pulse
    if s > 1.0: s = 1.0
    if s < -1.0: s = -1.0
    buf.append(int(s * 32767))

with wave.open(OUT, 'wb') as wf:
    wf.setnchannels(1)
    wf.setsampwidth(2)
    wf.setframerate(SR)
    wf.writeframes(buf.tobytes())

print(f"✅ 已生成: {OUT}")
print(f"   时长 {DUR}s | 采样率 {SR} | 单声道 16-bit | 大小 {len(buf)*2/1024:.0f} KB | 耗时 {time.time()-t0:.1f}s")
