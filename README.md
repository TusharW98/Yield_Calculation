
Yield to Maturity (YTM) Calculation using Excel VBA:

Overview:

This project demonstrates the calculation of Yield to Maturity (YTM) for a bond's dirty price using the Newton-Raphson method implemented in Excel VBA. 
The Newton-Raphson method is an iterative technique used to solve complex equations and has been applied here to determine the YTM with precision.

Features:

YTM Calculation: Computes the Yield to Maturity of a bond based on its dirty price.

Newton-Raphson Method: Uses iterative calculations to refine the YTM estimate until the bond price aligns with the market price.

Accuracy: Results are validated against Bloomberg Terminal data to ensure precision.

Implementation
VBA Code:

The VBA code is designed to implement the Newton-Raphson method for solving the YTM equation iteratively.
It includes a loop structure to adjust the YTM guess until the calculated bond price matches the market price.

Initial Guesses:
Initial guesses for YTM are explored based on financial principles to start the iterative process.

Precision and Efficiency:

The method achieves precise results with minimal iterations, demonstrating its effectiveness for bond valuation.

Review Results:

The calculated YTM will be displayed in the YTM used cell in the Excel sheet. Compare it with Bloomberg Terminal data for validation.
Results
The YTM calculated using this method matched the value from the Bloomberg Terminal, showcasing the accuracy and reliability of the Newton-Raphson approach.

Files

[Bond_Price.xlsx] - The Excel file containing the VBA code and data used for YTM calculations
