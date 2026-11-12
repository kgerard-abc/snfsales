"""Temporary: transcribe a Council meeting recording with timestamps.
usage: transcribe.py AUDIO OUT_PREFIX"""
import sys, json, time
from faster_whisper import WhisperModel

audio, out = sys.argv[1], sys.argv[2]
model = WhisperModel('small.en', device='cpu', compute_type='int8', cpu_threads=4)
segments, info = model.transcribe(
    audio, language='en', beam_size=2, vad_filter=True,
    vad_parameters={'min_silence_duration_ms': 700},
    initial_prompt=('Rhode Island Health Services Council meeting. Change in effective control. '
                    'AdviniaCare, Mayflower, Respiratory and Rehabilitation Center of Rhode Island, Coventry, '
                    'Greenville, Warren, RIDOH, Ombudsman, Medicaid, CMS, star rating, immediate jeopardy, '
                    'Attorney Rocha, Talamona.'))


def ts(s):
    s = int(s)
    return f'{s // 3600}:{s % 3600 // 60:02d}:{s % 60:02d}'


rows = []
t0 = time.time()
with open(out + '.txt', 'w') as f:
    for seg in segments:
        rows.append({'start': round(seg.start, 1), 'end': round(seg.end, 1), 'text': seg.text.strip()})
        f.write(f'[{ts(seg.start)}] {seg.text.strip()}\n')
        f.flush()
        if len(rows) % 100 == 0:
            print(ts(seg.end), 'done after', int(time.time() - t0), 's', flush=True)
json.dump({'duration': info.duration, 'segments': rows}, open(out + '.json', 'w'), indent=0)
print('finished', ts(info.duration))
