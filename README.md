# Portfolio Allocation Model

This document's purpose is to find the optimal allocation for any public equity portfolio with up to 30 assets based on the selected allocation method.

## Allocation Methods
There are 4 allocation methods available:
- **Equally Weighted (EW)**
- **Inverse Volatility (IV)**
- **Equal Risk Contribution (ERC)**
- **Minimal Volatility (MinVol)**

The model is dynamic: it adapts automatically to the number of assets inputted and the allocation method selected.  
To make it as such, i had relied heavily on `=IF()` and `=INDEX(MATCH(),MATCH())` functions. Which proved too complex for the Excel Solver in heavier calculations ("ERC"** and "MinVol" weights).  
This is why the weight distributions for "ERC" and "MinVol" are each solved for in their own Python scripts.  

---

## How to Use

1. **Security Entry Worksheet**  
   Input the tickers of all public equity securities you want in the portfolio. Bloomberg functions extract key market statistics on these stocks.

2. **PX Data Worksheet**  
   Retrieves the last 3 years of closing prices for the selected securities using Bloomberg. Daily performance of each stock over the period is calculated.

3. **GARCH Vol Worksheet**  
   Calculates each stock's volatility using the GARCH method.  
   - Two macro buttons are available in the upper-left corner:  
     1. Reset the α and β values of every security.  
     2. Copy the α and β values from the first asset to all others.

4. **Matrixes Worksheet**  
   Calculates and visualizes three matrices (useful for debugging and data analysis):  
   - The portfolio's **Correlation Matrix**  
   - Its **Variance-Covariance Matrix (Omega)**  
   - The **Weighted Omega Matrix** (calculated as `Omega * Weights`, dynamically adjusted based on the chosen allocation method).

5. **Portfolio Worksheet**  
   Outputs the weights of all inputted assets based on the selected allocation method (selectable via the "DistMeth" dropdown list in the upper-right corner).  
   - Calculates portfolio volatility and the Marginal Risk Contribution (%) of each asset.  
   - For **ERC** or **MinVol** methods:  
     - Click the macro button in the upper-left corner of the worksheet to execute the appropriate Python script (`Python_ERC_Solver` or `Python_MinVol_Solver`) for weight distribution calculations.

---

## Dependencies
- **Python Scripts**  
  - `Python_ERC_Solver.py`  
  - `Python_MinVol_Solver.py`  
- **Bloomberg Terminal Functions**  
- **Excel VBA Macros**  
  - To reset or apply α and β values.  
  - To run Python scripts for ERC and MinVol.

---

## Notes
- Ensure Python and the required libraries are installed and configured.
- Bloomberg Terminal must be properly set up for data retrieval.




# Portfolio Allocation Model

This document's purpose is to find the optimal allocation to any public equity portfolio with up to 30 assets based on the favoured allocation method. /n
There are 4 options:
 > Equally Weighted (EW).
 > Inverse Volatility (IV).
 > Equal Risk Contribution (ERC).
 > Minimal Volatility (MinVol).

The model is dynamic: it adapts automatically to the number of assets inputed and the allocation method selected.
To make it as such, I had to use many "=IF()" and "=INDEX(MATCH(),MATCH())" functions which made heavier calculations (Weights for "ERC" & "MinVol") too complex for the Excel solver.
This is why the weight distributions for "ERC" and "MinVol" are each solved for in their own Python scripts.  


1 - In the "Security Entry" worksheet, input the tickers of all the punlic equity securities you want in the portfolio. Bloomberg functions extract key market statistics on these stocks.
2 - The "PX Data" worksheet uses Bloomberg to retrieve the last 3 years of these securities’ prices at close, calculating the daily performance of each stock for the last 3 years. 
3 - The "GARCH Vol" worksheet calculates each stock's volatility using the GARCH method. The solver can be used to optimise the α and β values on the first security. 
    There are two macro buttons in the upper left corner: (1) for resetting the α and β values of every security, (2) for copying the values α and β from the 1st asset to all the others.
4 - The "Matrixes" worksheet calculates 3 matrices (mostly used for data visualisation and debugging): 
      > the portfolio's correlation matrix
      > its Variance Covariance matrix (Omega)
      > the Weighted Omega matrix. (Omega * Weights) 
    The last one depends on the choice of allocation method in the "Portfolio" worksheet. (Adjusts dynamically)
5 - The "Portfolio" worksheet yields the weights of all the inputted assets the method selected (in the upper-right corner "DistMeth" drop down list). 
    According to the selected method, the sheet calculates Portfolio Volatility, as well as the Marginal Risk Contribution (%) of every asset in the portfolio. 
    If the selected allocation methods are either "ERC" or "MinVol", you will need to click on the macro button in the upper left corner of the worksheet.
    The button is tied to some simple VBA code which will run the correct python scripts (either "Python_ERC_Solver" or "Python_MinVol_Solver"), to solve for each method's weight distributions.


## Directives 
In order to use this model, both python scripts MUST BE IN THE SAME FOLDER as this excel document, along with the "Settings" folder that goes with them. 
Macros must be activated, and automatic formulas must be turned on. 
Access to bloomberg is also needed, unless you have another way to input the closing prices of the past 750 working days for every security. 
If not, the prices and tickers already present on the document serve as a demo. 
There are notes scattered throughout the Excel Doc to help with potential errors.
