# DC Powerflow Model

Python implementation of the **DC Power Flow model** for electricity market clearing and pricing, with Excel-based input/output.

---

## Overview

The DC Power Flow model is a **linear approximation** of the AC power flow problem.  
It is widely used in electricity markets because it transforms a **non-convex optimization problem** into a **linear, convex problem** that can be solved efficiently.  

Key benefits:
- Linear constraints → fast and reliable optimization  
- Approximate solution for **active power flows**  
- Standard approach for **wholesale electricity market clearing in the U.S.**

---

## Key Assumptions

To achieve this linear approximation, the DC model makes the following simplifications:

- **Reactive power is ignored**  
- **Line resistances are assumed zero (𝑅ₗ = 0)** → active and reactive losses are neglected  
- **Voltage magnitudes are fixed at 1 p.u.** across all buses  
- **Phase angle differences between buses are assumed small** (𝜃ₖ − 𝜃ᵢ ≈ 0)  

These assumptions are accurate for **transmission networks**; in distribution networks, errors may be higher. For this project, the approximation is acceptable.

---

## Input Data

This repository does **not include the actual Excel input file**.  
Instead, the problem statement is provided as a screenshot:

`docs/dc_flow_problem_statement.png`

Users should create their own Excel input file following the structure in the statement.  
The script `src/dc_flow.py` will parse the file and produce outputs in Excel format.

---

## Outputs

The model generates the following outputs:

- **Social Welfare (€)**  
- **Congestion Surplus (€)**  
- **Optimal generation and demand allocation (MW)**  
- **Nodal prices (€/MWh)**  
- **Line power flows (MW)**  

All results are saved in `outputs/resultsdc1.xlsx`.

---

## Usage

1. Install required packages:
```bash
pip install -r requirements.txt
```
2.  Place your Excel input file in the appropriate folder.
3.  Run the script:
```bash
python src/dc_flow.py
```
4.  οpen `outputs/resultsdc.xlsx` to review results.

## Applications

* Electricity market clearing and pricing
* Transmission network congestion analysis
* Teaching and research in power systems and electricity markets

## Project Structure

dc-flow/
├── src/
│   └── dc_flow.py
├── outputs/
│   └── sample_output.xlsx
├── docs/
│   └── dc_flow_problem_statement.png
├── requirements.txt

## License

This project is licensed under the MIT License.
