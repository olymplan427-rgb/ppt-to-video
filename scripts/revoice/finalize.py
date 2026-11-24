# -*- coding: utf-8 -*-
"""마감(마스터링) 파이프라인 — 음량 정규화(-16 LUFS) · 오버레이 패널 · BGM 페이드/덕킹.

`revoice.build render`가 각 씬 세그먼트(`build/seg_<id>.mp4`)를 만든 뒤,
전체 씬의 음량을 -16 LUFS로 균일화하고(원음 클립·인터뷰·TTS 간 편차 제거),
필요한 씬에 이미지 오버레이 카드를 얹고, 광고 원음 구간 직전에 BGM을 페이드아웃해
최종 배포용 mp4를 완성한다.

CLI 사용 예:
  python -m revoice.finalize --manifest manifest.json --bgm bgm/track.wav --out build/final.mp4
  python -m revoice.finalize --manifest manifest.json --bgm bgm/track.wav \\
      --bgm-stop-before DP44A --copy-to ../../final.mp4
"""
import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys

from revoice import compose, env
from revoice.manifest import load_manifest

env.load_dotenv()

TARGET_LUFS = -16.0
LEAD = compose._LEAD_IN


def _load_font(weight: str, size: int):
    from PIL import ImageFont
    candidates = [
        os.path.join("assets", "fonts", f"WantedSans-{weight}.ttf"),
        os.path.join(os.path.dirname(os.path.dirname(__file__)), "assets", "fonts", f"WantedSans-{weight}.ttf"),
        "/System/Library/Fonts/AppleSDGothicNeo.ttc",
        r"C:\Windows\Fonts\malgunbd.ttf" if weight in ("Bold", "ExtraBold") else r"C:\Windows\Fonts\malgun.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except OSError:
                continue
    return ImageFont.load_default()


def render_overlay_panel(spec: dict, img_path: str, label: str, out_png: str) -> str:
    """라운드 카드 + 드롭섀도우 + 상단 필 라벨 오버레이 PNG(1920x1080 RGBA) 생성."""
    from PIL import Image, ImageDraw, ImageFilter
    x0, y0, x1, y1 = spec["box"]
    bw, bh = x1 - x0, y1 - y0
    fr = Image.new("RGBA", (1920, 1080), (0, 0, 0, 0))
    sh = Image.new("RGBA", (1920, 1080), (0, 0, 0, 0))
    ImageDraw.Draw(sh).rounded_rectangle((x0 + 6, y0 + 12, x1 + 6, y1 + 12), 28, fill=(0, 0, 0, 120))
    fr = Image.alpha_composite(fr, sh.filter(ImageFilter.GaussianBlur(14)))
    d = ImageDraw.Draw(fr)
    d.rounded_rectangle((x0, y0, x1, y1), 28, fill=tuple(spec["back"]) + (255,))
    im = Image.open(img_path).convert("RGB")
    pad, top = 26, 70
    r = min((bw - 2 * pad) / im.width, (bh - top - pad) / im.height)
    im = im.resize((int(im.width * r), int(im.height * r)), Image.LANCZOS)
    ix = x0 + (bw - im.width) // 2
    iy = y0 + top + (bh - top - pad - im.height) // 2
    m = Image.new("L", im.size, 0)
    ImageDraw.Draw(m).rounded_rectangle((0, 0, *im.size), 14, fill=255)
    fr.paste(im, (ix, iy), m)
    d.rounded_rectangle((ix, iy, ix + im.width, iy + im.height), 14, outline=(0, 0, 0, 40), width=2)
    font = _load_font("Bold", 30)
    tw = d.textlength(label, font=font)
    d.rounded_rectangle((x0 + 26, y0 + 18, x0 + 26 + tw + 40, y0 + 62), 22, fill=tuple(spec["pill"]) + (255,))
    d.text((x0 + 46, y0 + 40), label, font=font, fill=(255, 255, 255), anchor="lm")
    fr.save(out_png)
    return out_png


def _stamp(path: str) -> str:
    try:
        st = os.stat(path)
        return f"{st.st_size}:{int(st.st_mtime)}"
    except OSError:
        return "x"


def probe_duration(path: str) -> float:
    return float(subprocess.check_output(
        ["ffprobe", "-v", "error", "-show_entries", "format=duration",
         "-of", "csv=p=0", path]).decode().strip())


def measure_lufs(path: str):
    r = subprocess.run(
        ["ffmpeg", "-hide_banner", "-nostats", "-i", path, "-vn", "-af",
         "loudnorm=I=-16:TP=-1.5:LRA=11:print_format=json", "-f", "null", "-"],
        capture_output=True, text=True, encoding="utf-8", errors="replace")
    m = re.search(r'"input_i"\s*:\s*"(-?[\d.]+|-inf)"', r.stderr)
    return float(m.group(1)) if m and m.group(1) != "-inf" else None


def _load_cache(fin_dir: str) -> dict:
    try:
        with open(os.path.join(fin_dir, "_cache.json"), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def _save_cache(fin_dir: str, cache: dict) -> None:
    with open(os.path.join(fin_dir, "_cache.json"), "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False)


def process_scene(sc, seg: str, fin_dir: str, cache: dict, *,
                  target_lufs: float = TARGET_LUFS,
                  overlays: dict = None,
                  special_audio: dict = None):
    """단일 씬 세그먼트에 오버레이(선택)와 -16 LUFS 정적 게인 정규화를 적용."""
    overlays = overlays or {}
    special_audio = special_audio or {}
    out = os.path.join(fin_dir, f"fin_{sc.id}.mp4")
    ov = overlays.get(sc.id)
    sp = special_audio.get(sc.id)
    key = hashlib.sha1(json.dumps(
        [_stamp(seg), target_lufs, ov,
         [_stamp(i[0]) for i in (ov or {}).get("items", [])], sp]).encode()).hexdigest()[:12]

    if os.path.exists(out) and cache.get(sc.id) == key:
        return out, cache.get(sc.id + ":g")

    cmd = ["ffmpeg", "-y", "-v", "error", "-i", seg]
    vcodec = ["-c:v", "copy"]
    fc = None
    if ov:
        dur = probe_duration(seg)
        for i, (img, label, a, b) in enumerate(ov["items"]):
            pn = os.path.join(fin_dir, f"ov_{sc.id}_{i}.png")
            render_overlay_panel(ov, img, label, pn)
            cmd += ["-loop", "1", "-framerate", "30", "-t", f"{dur:.2f}", "-i", pn]
        chain, cur = [], "[0:v]"
        its = ov["items"]
        for i, (img, label, a, b) in enumerate(its):
            a += LEAD
            b += LEAD
            last = i == len(its) - 1
            end = b if last else its[i + 1][2] + LEAD + 0.35
            f = f"[{i+1}:v]format=rgba,fade=t=in:st={a:.2f}:d=0.18:alpha=1"
            if last:
                f += f",fade=t=out:st={b-0.18:.2f}:d=0.18:alpha=1"
            chain.append(f + f"[o{i}]")
            chain.append(
                f"{cur}[o{i}]overlay=0:0:eof_action=pass:"
                f"enable='between(t,{a:.2f},{end:.2f})':format=auto[v{i}]")
            cur = f"[v{i}]"
        fc = ";".join(chain)
        vcodec = ["-c:v", "libx264", "-preset", "medium", "-crf", "17",
                  "-pix_fmt", "yuv420p", "-r", compose.FPS]

    L = measure_lufs(seg)
    sp_cfg = sp or {}
    mg = sp_cfg.get("maxgain", 12.0)
    gain = 0.0 if L is None else max(-12.0, min(mg, sp_cfg.get("target", target_lufs) - L))
    af = sp_cfg.get("pre", "") + f"volume={gain:.2f}dB,alimiter=limit=0.89:level=false"
    if fc:
        cmd += ["-filter_complex", fc, "-map", f"[v{len(ov['items'])-1}]", "-map", "0:a"]
    else:
        cmd += ["-map", "0:v", "-map", "0:a"]
    cmd += vcodec + ["-af", af, "-c:a", "aac", "-ar", "48000", "-ac", "2", "-b:a", "192k", out]
    subprocess.run(cmd, check=True)
    cache[sc.id] = key
    cache[sc.id + ":g"] = [L, gain]
    _save_cache(fin_dir, cache)
    return out, [L, gain]


def finalize_manifest(manifest_path: str, *,
                      build_dir: str = "build",
                      bgm_path: str = "",
                      out_path: str = "",
                      copy_to: str = "",
                      overlays: dict = None,
                      special_audio: dict = None,
                      bgm_stop_before: str = "",
                      bgm_gain_db: float = -18.0,
                      target_lufs: float = TARGET_LUFS,
                      highlight_suffix: str = "") -> str:
    """매니페스트의 모든 세그먼트를 정규화·합본·BGM 덕킹해 최종 mp4 생성."""
    m = load_manifest(manifest_path)
    fin_dir = os.path.join(build_dir, "fin")
    os.makedirs(fin_dir, exist_ok=True)
    cache = _load_cache(fin_dir)

    paths, rows = [], []
    for sc in m.scenes:
        seg = os.path.join(build_dir, f"seg_{sc.id}.mp4")
        if not os.path.exists(seg):
            raise SystemExit(f"세그먼트 누락: {seg} — 먼저 `python -m revoice.build render --manifest {manifest_path}`를 실행하세요.")
        out_sc, info = process_scene(
            sc, seg, fin_dir, cache,
            target_lufs=target_lufs, overlays=overlays, special_audio=special_audio)
        paths.append(out_sc)
        rows.append((sc.id, info))

    print("씬        측정LUFS → 게인(dB)")
    for sid, info in rows:
        if info and info[0] is not None:
            mark = " *" if (highlight_suffix and sid.endswith(highlight_suffix)) else ""
            print(f"  {sid:10s} {info[0]:6.1f} → {info[1]:+5.1f}{mark}")

    out_path = out_path or m.output
    prefix = os.path.splitext(os.path.basename(out_path))[0]
    lst = os.path.join(fin_dir, f"_list_{prefix}.txt")
    with open(lst, "w", encoding="utf-8") as f:
        for p in paths:
            f.write(f"file '{os.path.basename(p)}'\n")

    base = os.path.join(fin_dir, prefix + "_base.mp4")
    compose.run_ffmpeg(compose.build_concat_cmd(lst, base))

    bgm_file = bgm_path or os.path.join("bgm", f"{prefix}.wav")
    if bgm_file and os.path.exists(bgm_file):
        bgm_stop = None
        if bgm_stop_before:
            acc = 0.0
            for sc, p in zip(m.scenes, paths):
                if sc.id == bgm_stop_before:
                    bgm_stop = acc
                    break
                acc += probe_duration(p)

        if bgm_stop is not None:
            fade_st = max(0.0, bgm_stop - 1.5)
            fc = (
                f"[1:a]volume={bgm_gain_db:.1f}dB,aresample=48000,"
                f"afade=t=out:st={fade_st:.2f}:d=1.50,"
                f"volume=enable='gte(t,{bgm_stop:.2f})':volume=0[bg];"
                f"[bg][0:a]sidechaincompress=threshold=0.05:ratio=8:attack=20:release=300[duck];"
                f"[0:a][duck]amix=inputs=2:duration=first:normalize=0[a]"
            )
            compose.run_ffmpeg([
                "ffmpeg", "-y",
                "-i", base,
                "-stream_loop", "-1", "-i", bgm_file,
                "-filter_complex", fc,
                "-map", "0:v", "-map", "[a]",
                "-c:v", "copy", "-c:a", "aac", "-ar", "48000", "-ac", "2",
                "-shortest", out_path,
            ])
            print(f"[bgm]  설명회 BGM {bgm_stop:.2f}s 직전 페이드아웃 ({bgm_stop_before} 원음 보호)")
        else:
            compose.run_ffmpeg(compose.build_bgm_duck_cmd(base, bgm_file, out_path, bgm_gain_db=bgm_gain_db))
    else:
        shutil.copy2(base, out_path)
        print("[bgm]  BGM 파일 없음 — 무BGM 마스터로 저장")

    if copy_to:
        os.makedirs(os.path.dirname(os.path.abspath(copy_to)), exist_ok=True)
        shutil.copy2(out_path, copy_to)
        print(f"[copy] {copy_to}")

    dur = probe_duration(out_path)
    print(f"[done] {out_path} ({int(dur // 60):02d}:{dur % 60:05.2f})")
    return out_path


def main(argv=None):
    ap = argparse.ArgumentParser(description="씬 음량 정규화(-16 LUFS) · 오버레이 · BGM 마스터링")
    ap.add_argument("--manifest", required=True, help="매니페스트 JSON 경로")
    ap.add_argument("--bgm", default="", help="BGM WAV 경로 (생략 시 bgm/<output_stem>.wav)")
    ap.add_argument("--out", default="", help="최종 출력 mp4 경로 (생략 시 manifest.output)")
    ap.add_argument("--copy-to", default="", help="추가 복사 경로(예: 프로젝트 루트 최종본)")
    ap.add_argument("--bgm-stop-before", default="", help="이 씬 ID 직전에 BGM을 1.5초 페이드아웃 후 음소거")
    ap.add_argument("--lufs", type=float, default=TARGET_LUFS, help="목표 LUFS (기본 -16.0)")
    args = ap.parse_args(argv)
    finalize_manifest(
        args.manifest,
        bgm_path=args.bgm,
        out_path=args.out,
        copy_to=args.copy_to,
        bgm_stop_before=args.bgm_stop_before,
        target_lufs=args.lufs,
    )


if __name__ == "__main__":
    main()
