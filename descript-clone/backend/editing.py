import subprocess


def merge_ranges(ranges, gap=0.0):
    """Ordena y une rangos [start, end] que se tocan (o distan <= gap)."""
    out = []
    for s, e in sorted((float(s), float(e)) for s, e in ranges if e > s):
        if out and s - out[-1][1] <= gap:
            out[-1][1] = max(out[-1][1], e)
        else:
            out.append([s, e])
    return out


def has_video(path: str) -> bool:
    r = subprocess.run(
        ["ffprobe", "-v", "error", "-select_streams", "v", "-show_entries",
         "stream=codec_type", "-of", "csv=p=0", path],
        capture_output=True, text=True,
    )
    return "video" in r.stdout


def build_command(src: str, dst: str, keep, video: bool) -> list[str]:
    """Comando ffmpeg que concatena solo los rangos `keep` (re-encode, cortes precisos)."""
    parts, labels = [], []
    for i, (s, e) in enumerate(keep):
        if video:
            parts.append(f"[0:v]trim=start={s}:end={e},setpts=PTS-STARTPTS[v{i}]")
            labels.append(f"[v{i}]")
        parts.append(f"[0:a]atrim=start={s}:end={e},asetpts=PTS-STARTPTS[a{i}]")
        labels.append(f"[a{i}]")
    n = len(keep)
    concat = f"{''.join(labels)}concat=n={n}:v={1 if video else 0}:a=1" + ("[v][a]" if video else "[a]")
    cmd = ["ffmpeg", "-y", "-i", src, "-filter_complex", ";".join(parts + [concat])]
    cmd += ["-map", "[v]"] if video else []
    cmd += ["-map", "[a]"]
    cmd += ["-c:v", "libx264", "-preset", "veryfast", "-pix_fmt", "yuv420p"] if video else []
    cmd += ["-c:a", "aac" if video else "libmp3lame", dst]
    return cmd


def export(src: str, dst: str, keep) -> None:
    keep = merge_ranges(keep)
    if not keep:
        raise ValueError("No queda nada para exportar")
    r = subprocess.run(build_command(src, dst, keep, has_video(src)), capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError(r.stderr[-2000:])
