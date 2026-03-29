# PathFinder Tool: User Guide

Welcome to the **PathFinder** user guide.
This software is designed to help industrial sites (such as refineries) define their **optimal decarbonization strategy** from 2025 to 2050.

In concrete terms, the tool analyzes all available green technologies, your budget constraints, and your CO2 reduction targets to mathematically find **the cheapest investment plan** that complies with all your rules.

---

## 1. Prerequisites and Installation

To use the tool on your computer, you need:
1. **Python 3.9+** installed on your machine.
2. Install the required libraries by opening a terminal in this folder and running:
   ```bash
   pip install pandas pulp matplotlib openpyxl xlsxwriter
   ```

---

## 2. How it works (Overview)

Using the tool is done in **3 simple steps**:
1. You **fill in your data** in the main Excel file (`PathFinder input.xlsx`).
2. You **run the calculation** via the Python script.
3. You **analyze the results** (charts and tables) automatically generated in an output folder.

---

## 3. Step 1: Set up the Input Excel File

All the configuration is done in the **`PathFinder input.xlsx`** file. This file is the "brain" of your data. The tool will read this information to understand your factory.

**Main tabs to fill out:**
- **Time Series**: Here you indicate the price forecasts for electricity, gas, hydrogen, and the cost of the carbon tax, year by year.
- **Technologies**: The list of machines or upgrades you can buy (e.g., electric boiler, electrolyzer). Specify their cost (CAPEX), their maintenance (OPEX), and the date from which they become available on the market.
- **Process**: The baseline of your factory. It describes how much each part of the factory currently consumes.
- **Loans & CCfD**: Configure the bank borrowing rates and potential state subsidies.

⚠️ **Golden Rule:** Do not delete the `[START]` and `[END]` tags in the Excel file. The program uses them to read tables automatically, even if you add new rows.

---

## 4. Step 2: Run the Optimization

Once your Excel file is securely saved and closed, you are ready to launch the mathematical calculation.

Open a terminal (command prompt) in the project folder, and simply type:
```bash
python main.py
```

**What happens next?**
The computer will start processing ("Solving solver..."). If there are many choices to make for the factory, this can take anywhere from a few seconds to a few minutes. The solver checks all possible combinations to find the financial optimum.

---

## 5. Step 3: Analyze the Results

If the solver has found a mathematically viable solution (Status: `Optimal`), it will create new files containing your roadmap.

### A. The Excel Report (Output)
The tool will generate a detailed, row-by-row Excel file. It contains a complete summary of the solution chosen by the machine, year by year:
- How much energy you consume as the technologies are deployed.
- Your exact carbon footprint.
- Your total costs (Energy bills, loan repayments, carbon penalties).

### B. The Charts (Visual Dashboards)
The script will create a visual dashboard in PDF/Image format containing several key charts:
1. **Emissions Trajectory**: See how your direct and indirect emissions (Scope 1 & 2) decrease over time compared to your goals.
2. **Investment Plan (Roadmap)**: A Gantt chart or bar chart that shows you *which* technology to buy, and in *which year*.
3. **Cost Breakdown**: Visualize exactly where the money goes (OPEX vs. Carbon Tax).
4. **Energy Mix**: Discover how your consumption shifts from gas/fuel to electricity and green hydrogen.

---

## 6. Common Issues (FAQ)

### "The program shows an Infeasible error"
**What it means:** The computer found **no** possible mathematical solution that respects all your constraints simultaneously. 
**How to solve it:** 
- Have you set a maximum budget too low to reach overly ambitious CO2 targets?
- Try increasing the budget, or softening the reduction targets ("Objectives" in the Excel file) to see if the calculation passes.

### "The computer proposes no new construction"
**What it means:** The solver calculated that it is mathematically more profitable (cheaper) for your company to keep the old machines and pay the carbon tax, rather than buying expensive green technologies.
**How to solve it:** 
- Check if the carbon tax price is high enough in your forecasts to "force" the transition.
- Add subsidies or reduce the initial cost (CAPEX) of your green technologies to make them competitive.

---
*This guide focuses on practical use. To understand the exact mathematical mechanics (MILP equations, integral variables), please refer to the `README_Optimization_Logic.md` document.*
