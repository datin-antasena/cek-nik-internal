from datetime import datetime
from zoneinfo import ZoneInfo


def catat_log(nama_file, nama_sheet, rincian_per_kolom):
    waktu = datetime.now(ZoneInfo("Asia/Jakarta")).strftime("%Y-%m-%d %H:%M:%S")
    summary_text = ""

    for col, stats in rincian_per_kolom.items():
        simple_stats = {}
        ganda_total = 0
        for k, v in stats.items():
            if str(k).startswith("GANDA"):
                ganda_total += v
            else:
                simple_stats[k] = v
        if ganda_total > 0:
            simple_stats["GANDA (TOTAL)"] = ganda_total

        stat_str = ", ".join(f"{k}:{v}" for k, v in simple_stats.items())
        summary_text += f"[{col}: {stat_str}] "

    pesan = (
        f"[{waktu}] FILE: {nama_file} | SHEET: {nama_sheet} | DETAIL: {summary_text}\n"
    )
    with open("activity_log.txt", "a") as f:
        f.write(pesan)
