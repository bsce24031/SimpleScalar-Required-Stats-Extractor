#!/usr/bin/env python3
"""
SimpleScalar Required Stats Extractor
=======================================================
Developed by BSCE-24031 for COA/CA Project (Spring 2026) — ITU Lahore

Extracts and displays ONLY the required statistics:
  mem_size
  sim_cycle, sim_num_insn, sim_num_refs
  sim_num_loads, sim_num_stores, sim_num_branches
  sim_inst_rate, sim_CPI, ruu_occupancy
  For ALL caches: *.accesses, *.hits, *.miss_rate

Output: terminal only — no files saved.
"""

import re
import os

# ─────────────────────────────────────────────────────────────────────────────
#  STATS TO EXTRACT
# ─────────────────────────────────────────────────────────────────────────────

CORE_STATS = [
    ("mem_size",         r"mem_size\s+([\d.e+]+)",         "Memory Size"),
    ("sim_cycle",        r"sim_cycle\s+([\d.e+]+)",         "Sim Cycles"),
    ("sim_num_insn",     r"sim_num_insn\s+([\d.e+]+)",      "Instructions"),
    ("sim_num_refs",     r"sim_num_refs\s+([\d.e+]+)",      "Memory References"),
    ("sim_num_loads",    r"sim_num_loads\s+([\d.e+]+)",     "Loads"),
    ("sim_num_stores",   r"sim_num_stores\s+([\d.e+]+)",    "Stores"),
    ("sim_num_branches", r"sim_num_branches\s+([\d.e+]+)",  "Branches"),
    ("sim_inst_rate",    r"sim_inst_rate\s+([\d.e+]+)",     "Inst Rate"),
    ("sim_CPI",          r"sim_CPI\s+([\d.e+]+)",           "CPI"),
    ("ruu_occupancy",    r"ruu_occupancy\s+([\d.e+]+)",     "RUU Occupancy"),
]

CACHE_NAMES = ["il1", "dl1", "il2", "dl2", "ul2"]

# ─────────────────────────────────────────────────────────────────────────────
#  PARSER
# ─────────────────────────────────────────────────────────────────────────────

def parse_file(filepath):
    with open(filepath, "r", errors="replace") as f:
        text = f.read()

    stats = {}

    for key, pattern, _ in CORE_STATS:
        m = re.search(pattern, text, re.MULTILINE)
        if m:
            stats[key] = float(m.group(1))

    # Auto-detect all caches present in file
    found_caches = []
    for cname in CACHE_NAMES:
        if re.search(rf"^{cname}\.accesses\s+", text, re.MULTILINE):
            found_caches.append(cname)
    for cname in re.findall(r"^(\w+)\.accesses\s+[\d.e+]+", text, re.MULTILINE):
        if cname not in found_caches:
            found_caches.append(cname)

    stats["_caches"] = found_caches

    for cname in found_caches:
        for metric in ("accesses", "hits", "miss_rate"):
            m = re.search(rf"^{re.escape(cname)}\.{metric}\s+([\d.e+]+)",
                          text, re.MULTILINE)
            if m:
                stats[f"{cname}.{metric}"] = float(m.group(1))

    return stats


def all_caches(stats_list):
    seen = []
    for s in stats_list:
        for c in s.get("_caches", []):
            if c not in seen:
                seen.append(c)
    return seen


# ─────────────────────────────────────────────────────────────────────────────
#  FORMATTER
# ─────────────────────────────────────────────────────────────────────────────

def fmt(value, key=""):
    if value is None:
        return "N/A"
    if "miss_rate" in key:
        return f"{value:.6f}"
    if key in ("sim_CPI", "ruu_occupancy", "sim_inst_rate"):
        return f"{value:.4f}"
    v = float(value)
    if v >= 1e9:  return f"{v/1e9:.4f} G"
    if v >= 1e6:  return f"{v/1e6:.4f} M"
    if v >= 1e3:  return f"{v:,.0f}"
    return f"{v:.4f}"


# ─────────────────────────────────────────────────────────────────────────────
#  TERMINAL TABLE PRINTER
# ─────────────────────────────────────────────────────────────────────────────

def print_table(benchmark_name, config_labels, stats_list):
    caches = all_caches(stats_list)
    n      = len(config_labels)
    LABEL  = 26
    COL    = 16
    total_w = LABEL + (COL + 1) * n

    def section(title):
        print(f"╟{'─'*LABEL}{'┼' + '┼'.join('─'*COL for _ in config_labels)}╢")
        print(f"║ ── {title:<{total_w - 5}}║")
        print(f"╟{'─'*LABEL}{'┼' + '┼'.join('─'*COL for _ in config_labels)}╢")

    def data_row(label, vals):
        cells = "│".join(f"{v:>{COL}}" for v in vals)
        print(f"│ {label:<{LABEL}}{cells}│")

    bar = "═" * total_w
    print(f"\n╔{bar}╗")
    print(f"║{f' Benchmark: {benchmark_name} ':^{total_w}}║")
    print(f"╠{bar}╣")
    hdr = f"{'Metric':<{LABEL}}" + "│".join(f"{lbl:^{COL}}" for lbl in config_labels)
    print(f"║ {hdr}║")
    print(f"╠{'═'*LABEL}{'╪' + '╪'.join('═'*COL for _ in config_labels)}╣")

    section("Core Statistics")
    for key, _, label in CORE_STATS:
        data_row(f"  {label}", [fmt(s.get(key), key) for s in stats_list])

    for cname in caches:
        section(f"Cache: {cname.upper()}")
        for metric in ("accesses", "hits", "miss_rate"):
            fk = f"{cname}.{metric}"
            data_row(f"  {fk}", [fmt(s.get(fk), fk) for s in stats_list])

    print(f"╚{bar}╝")


# ─────────────────────────────────────────────────────────────────────────────
#  INTERACTIVE INPUT
# ─────────────────────────────────────────────────────────────────────────────

CONFIG_DEFAULTS = ["Small", "Medium", "Large", "Excessive"]

def prompt_benchmark(idx, total):
    print(f"\n{'═'*56}")
    print(f"  BENCHMARK {idx} of {total}")
    print(f"{'═'*56}")
    bname = input("  Benchmark name (e.g. ijpeg): ").strip() or f"Benchmark{idx}"

    print(f"\n  Enter result files for '{bname}'.")
    print("  Press ENTER on an empty path when done.\n")

    configs = []
    while True:
        n = len(configs)
        default = CONFIG_DEFAULTS[n] if n < len(CONFIG_DEFAULTS) else f"Config{n+1}"
        path = input(f"  File {n+1} path (ENTER to finish): ").strip()
        if not path:
            if not configs:
                print("  Please enter at least one file.")
                continue
            break
        if not os.path.isfile(path):
            print(f"  ✗ File not found: {path}")
            continue
        label = input(f"  Label [{default}]: ").strip() or default
        stats = parse_file(path)
        core_found  = sum(1 for k, *_ in CORE_STATS if k in stats)
        cache_found = sum(1 for k in stats if "." in k)
        print(f"  ✓  {os.path.basename(path)}: {core_found} core stats, {cache_found} cache values")
        configs.append((label, stats))

    return bname, [c[0] for c in configs], [c[1] for c in configs]


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("\n" + "="*56)
    print("  SimpleScalar Required Stats — Terminal Lite")
    print("  COA/CA Project — Spring 2026 | ITU Lahore")
    print("="*56)

    while True:
        nb = input("\n  How many benchmarks? [1]: ").strip() or "1"
        try:
            nb = int(nb)
            if nb >= 1:
                break
        except ValueError:
            pass
        print("  Enter a positive integer.")

    all_data = []
    for i in range(1, nb + 1):
        bname, labels, stats_list = prompt_benchmark(i, nb)
        all_data.append((bname, labels, stats_list))

    print("\n" + "="*56)
    print("  RESULTS")
    print("="*56)

    for bname, labels, stats_list in all_data:
        print_table(bname, labels, stats_list)


if __name__ == "__main__":
    main()
