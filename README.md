### Trading.xlsm - Market Simulation and Option Pricing Tool 

This Excel workbook, Trading.xlsm, is a quantitative tool built with VBA macros. It is designed to demonstrate a real-time market simulation and to perform numerical option pricing by integrating the payoff function over the underlying asset's probability distribution.

### Key Features
Real-time Market Simulation: The workbook includes a custom market simulation engine that generates a dynamic spot price (SpotMidMarket). The simulation can be started, paused, and reset, with all price movements logged for analysis.

Numerical Option Pricing: Implements a method for pricing a European-style call option by numerically integrating the payoff function over a discretized normal probability distribution of the underlying asset.

Payoff Calculation & Visualization: The tool automatically calculates and displays the payoff function (Max(Spot - Strike, 0)) for a given strike price.

Probability Distribution Visualization: It simulates and generates a plot of the probability density function (PDF) of the spot price at a future date, providing a clear visual of the potential outcomes.

Dynamic Charting: Automatically generates and updates scatter plots to visualize both the market simulation history and the option pricing results (payoff and probability distribution).

### VBA Macros
The following VBA macros power the workbook's functionality:

GoButton / StopButton: Controls the state of the market simulation, allowing it to start, pause, and reset the spot price and recorded data.

MarketTick: The core simulation macro that updates the spot price with a random movement and logs the new data point at a defined time interval.

MonteCarloSpotDistribution: Calculates and populates the data tables for the option's payoff and the spot price's probability distribution, then generates the corresponding charts.

InsertScatterChart: Generates a scatter plot of the market simulation data (step vs. spot price).

### Prerequisites and Usage
This document requires macros to be enabled in Microsoft Excel to function correctly.

To use this tool:

Open the file Trading.xlsm.

If a security warning appears, be sure to enable content to allow macros to run.

Use the Go / Pause button to start the market simulation.

Modify the input parameters for the option pricing (e.g., strike price, volatility) to perform a new calculation.

### Author
Maïssane Frikh - Mathematical Engineering Student, with a strong interest in quantitative finance and numerical methods.
