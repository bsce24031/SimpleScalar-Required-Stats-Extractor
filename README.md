SimpleScalar Results Table Generator
<div align="center">
Automated Performance Analysis Tool for SimpleScalar
COA/CA Project — Spring 2026 | ITU Lahore

Developed By Muhammad Raza Mukhtar

</div>
 Project Overview

The SimpleScalar Results Table Generator is a Python-based automated analysis and reporting tool designed to process SimpleScalar sim-outorder simulation output files.

The tool automatically extracts important processor performance statistics, organizes them into structured comparison tables, and generates professionally formatted Excel reports for easy analysis.

It is specially designed for Computer Organization & Architecture (COA/CA) projects to simplify benchmark evaluation and performance comparison across different processor configurations.

The program supports:

Multiple benchmarks
Multiple processor configurations
Automatic regex-based stat extraction
Cache and memory analysis
Branch predictor analysis
Pipeline behavior analysis
Cross-benchmark comparison
Summary and findings generation
Formatted Excel report creation

The generated Excel workbook contains:

Detailed benchmark result sheets
Cross-benchmark comparison tables
Final results summary
Key outputs table
Important findings and insights
Professional formatting with colors and grouped sections

The tool also computes additional derived metrics such as:

Cache hit rates
Branch prediction miss rates
IPC ranges
Performance speedups
Cache miss rate comparisons
Cycle overhead analysis

This project eliminates the need to manually scan lengthy SimpleScalar output logs and helps students and researchers quickly analyze CPU simulation results.

 Features
✔ Automatic parsing of sim-outorder output files
✔ Supports multiple benchmarks and configurations
✔ Professional terminal tables
✔ Automatic Excel report generation
✔ Cross-benchmark comparison support
✔ Summary and findings generation
✔ Cache performance analysis
✔ Pipeline and branch predictor analysis
✔ Dynamic statistic extraction using regex
✔ Lightweight and easy to use
✔ Organized grouped statistics
✔ Human-readable formatted outputs
📂 Supported Benchmarks

Examples include:

ijpeg
gcc
compress
go
li
vortex
Any benchmark producing valid SimpleScalar output
Extracted Statistics
🖥 Core Performance Statistics
Total Instructions
Memory References
Load Instructions
Store Instructions
Branch Instructions
Total Cycles
IPC (Instructions Per Cycle)
CPI (Cycles Per Instruction)
Simulation Speed
Execution Bandwidth
Instructions per Branch
Average Simulation Slip
 Pipeline Statistics
IFQ (Instruction Fetch Queue)
IFQ Occupancy
IFQ Rate
IFQ Latency
IFQ Full Fraction
RUU / LSQ
RUU Occupancy
RUU Full Fraction
LSQ Occupancy
LSQ Full Fraction
🌿 Branch Predictor Statistics
Predictor Lookups
Predictor Updates
Address Hits
Direction Hits
Predictor Misses
Branch Direction Rate
Branch Address Rate
Jump Register Rate
💾 Cache Statistics
DL1 Cache
Accesses
Hits
Misses
Miss Rate
Replacement Rate
Writeback Rate
IL1 Cache
Accesses
Hits
Misses
Miss Rate
DL2 Cache
Accesses
Hits
Misses
Miss Rate
IL2 Cache
Accesses
Hits
Misses
Miss Rate
🧭 TLB & Memory Statistics
ITLB Miss Rate
DTLB Miss Rate
Memory Page Count
⚙ Configuration Parameters
IFQ Size
Decode Width
Issue Width
Commit Width
RUU Size
LSQ Size
Generated Excel Sheets

The tool automatically creates a formatted Excel workbook containing:

Benchmark Sheets

Detailed tables for each benchmark including:

Core performance
Cache statistics
Pipeline statistics
Branch predictor statistics
Memory statistics
Configuration parameters
📄 Cross-Benchmark Comparison Sheet

Compares important metrics across benchmarks including:

IPC
CPI
Total cycles
Cache miss rates
Branch prediction performance
Pipeline occupancy
Final Results Sheet

Displays:

Best IPC
Best CPI
Minimum cycles
Lowest cache miss rates
Best branch prediction rates
Key Outputs Sheet

Compact overview of:

IPC
CPI
Cycles
Cache miss rates
Branch prediction rates
RUU occupancy

for every benchmark configuration.

Important Findings Sheet

Automatically generated insights including:

Best IPC configuration
Worst IPC configuration
IPC speedup
Cache miss rate ranges
Branch predictor rate ranges
Cycle overhead analysis
🛠 Requirements
Required
Python 3.x
Optional (for Excel generation)

Install openpyxl:

pip install openpyxl
▶️ How to Run
python3 simscalar_tables.py
🧾 Program Workflow
Run the script
Enter number of benchmarks
Provide benchmark names
Enter simulation result file paths
Assign configuration labels
View terminal comparison tables
Automatically generate formatted Excel report
📂 Example Directory Structure
project/
│
├── simscalar_tables.py
├── outputs/
│   ├── ijpeg_small.txt
│   ├── ijpeg_medium.txt
│   ├── ijpeg_large.txt
│   └── gcc_small.txt
│
└── README.md
📋 Example Input
Benchmark name: ijpeg

File 1: ijpeg_small.txt
Label: Small

File 2: ijpeg_medium.txt
Label: Medium

File 3: ijpeg_large.txt
Label: Large
Example Terminal Output
Benchmark: ijpeg

Metric                     Small      Medium      Large
----------------------------------------------------------
Total Cycles               1.23M      2.45M      5.67M
IPC                        1.2345     1.4567     1.7890
CPI                        0.8123     0.6871     0.5589
DL1 Miss Rate              0.0234     0.0187     0.0151
RUU Occupancy              12.34      15.67      19.22
📁 Example Generated Output
tables_output/
└── simscalar_results_ijpeg.xlsx
Use Cases

This project is useful for:

Computer Architecture projects
CPU performance analysis
Benchmark comparison
Cache behavior analysis
Pipeline evaluation
Academic benchmarking reports
Research projects using SimpleScalar
⚡ Advantages
Saves time during result analysis
Eliminates manual parsing errors
Generates professional reports automatically
Easy to extend and customize
Beginner-friendly implementation
Lightweight and efficient
👨‍💻 Author

Muhammad Raza Mukhtar
COA/CA Project — Spring 2026
Information Technology University (ITU) Lahore

📌 Note

This tool does not modify simulation output files.

It only reads and extracts required statistics from SimpleScalar outputs for analysis, comparison, and report generation purposes.

📜 License

This project is developed for academic and educational purposes.
