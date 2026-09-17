# -*- coding: utf-8 -*-
import subprocess, os, shutil
import imageio_ffmpeg

FFMPEG = imageio_ffmpeg.get_ffmpeg_exe()
INPUT = "Demonstration_PFE_Toutes_Videos.mp4"
OUTPUT = "Demo_soutenance_final.mp4"
TEMP_DIR = "temp_video_parts"
os.makedirs(TEMP_DIR, exist_ok=True)

segments = [
    (0,    45,   1.0),
    (45,   145,  1.5),
    (145,  171,  1.0),
    (171,  360,  1.5),
    (360,  381,  1.0),
    (381,  413,  1.5),
    (413,  480,  1.0),
]

parts = []
for i, (start, end, speed) in enumerate(segments):
    part_file = os.path.join(TEMP_DIR, f"part_{i:02d}.mp4")
    parts.append(part_file)
    duration = end - start
    if speed == 1.0:
        cmd = [FFMPEG,"-y","-ss",str(start),"-t",str(duration),"-i",INPUT,"-c:v","libx264","-preset","fast","-c:a","aac",part_file]
    else:
        cmd = [FFMPEG,"-y","-ss",str(start),"-t",str(duration),"-i",INPUT,"-vf",f"setpts={1.0/speed}*PTS","-af",f"atempo={speed}","-c:v","libx264","-preset","fast","-c:a","aac",part_file]
    print(f"[{i+1}/{len(segments)}] {start}s-{end}s x{speed}")
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0: print("ERR:", r.stderr[-300:])
    else: print(f"  OK {os.path.getsize(part_file)//1024}Ko")

concat_file = os.path.join(TEMP_DIR, "concat.txt")
with open(concat_file,"w") as f:
    for p in parts: f.write(f"file '{os.path.abspath(p)}'\n")

cmd2=[FFMPEG,"-y","-f","concat","-safe","0","-i",concat_file,"-c:v","libx264","-preset","fast","-c:a","aac","-movflags","+faststart",OUTPUT]
print("Concatenation...")
r2=subprocess.run(cmd2,capture_output=True,text=True)
if r2.returncode!=0: print("ERR:",r2.stderr[-500:])
else: print(f"OK! {OUTPUT} ({os.path.getsize(OUTPUT)/1024/1024:.1f}Mo)")
shutil.rmtree(TEMP_DIR,ignore_errors=True)
