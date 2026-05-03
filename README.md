# SimpleScalar Required Stats Extractor — Terminal Lite

##  Project Title

**SimpleScalar Required Stats Extractor (Terminal Lite)**

---

##  GitHub Repository Description

A lightweight Python-based terminal tool to automatically extract and compare key performance statistics from SimpleScalar simulation outputs. Designed for Computer Architecture (COA/CA) projects, it helps students and researchers quickly analyze benchmark results across multiple configurations without manually scanning large output logs.

---

##  Overview

This tool parses SimpleScalar output files and extracts only the **required performance metrics**, presenting them in a clean terminal-based comparison table.

It supports:

* Multiple benchmarks (e.g., ijpeg, compress, gcc)
* Multiple configurations (Small, Medium, Large, etc.)
* Automatic cache detection
* Clean formatted CLI output

---

##  Extracted Metrics

### Core Statistics

* Memory size
* Simulation cycles
* Instruction count
* Memory references
* Loads / Stores / Branches
* Instruction rate
* CPI (Cycles Per Instruction)
* RUU occupancy

### Cache Metrics (Auto-detected)

For all detected caches (il1, dl1, il2, dl2, ul2, etc.):

* Accesses
* Hits
* Miss rate

---

##  Features

* ✔ Fully automated regex-based parsing
* ✔ No external files generated (terminal-only output)
* ✔ Supports multiple benchmark comparisons
* ✔ Dynamic cache detection
* ✔ Clean ASCII-style comparison tables
* ✔ Lightweight and fast execution

---

##  How to Run

```bash
python3 simscaler_stats_extractor.py
```

---

## 🧾 Usage Flow

1. Run the script
2. Enter number of benchmarks
3. Provide benchmark name (e.g., ijpeg)
4. Enter file paths of simulation outputs
5. Assign labels (Small, Medium, Large, etc.)
6. View formatted comparison table in terminal

---

## 📂 Example Output

```
Benchmark: ijpeg
Metric                Small        Medium       Large
------------------------------------------------------
Sim Cycles           1.23M        2.45M        5.67M
CPI                  1.2345       1.4567       1.7890
L1 Cache Hits        98.5%        97.2%        95.1%
```

---

## 🛠 Requirements

* Python 3.x
* No external libraries required (uses only built-in modules)

---

## 🎯 Use Case

* Computer Architecture (COA/CA) projects
* SimpleScalar result analysis
* Performance comparison across CPU configurations
* Academic benchmarking reports

---

## 👨‍🎓Author

Developed By Muhammad Raza Mukhtar for COA/CA Project (Spring 2026) — ITU Lahore

---

##  Note

This tool does not modify simulation outputs. It only reads and extracts required statistics for analysis purposes.
