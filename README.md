# Charging Strategy Planning Model (CSPM) for EVs
 
This repository contains the implementation of a novel Mixed-Integer Linear Programming (MILP) model designed to determine an optimal charging strategy for Electric Vehicles (EVs). The model aims to minimize travel time and cost for the Electric Vehicle Charging Strategy Problem (EVCSP) in a single-stage, providing insights into when, where, and how much charge is needed. This model accounts for multiple refueling stops, partial charging, nonlinear charging times, and varying power rates.

## Table of Contents

- [Problem Definition](#problem-definition)
- [Model Contributions](#model-contributions)
- [Installation](#installation)
- [Usage](#usage)
- [Example](#example)
- [Requirements](#requirements)
- [References](#references)
- [License](#license)

## Problem Definition

Given a route with an origin and a destination node, the EVCSP determines the optimal charging decisions, including:

- Charging location
- Charging duration
- Charge amount

The goal is to minimize travel time, cost, or both. 

### Assumptions

1. The number of charging stations (`n`) on the route is adequate to complete the trip.
2. The energy required to reach consecutive charging stations does not exceed the EV range.
3. The State of Charge (SOC) is not allowed to drop under a given threshold (`B_min`).
4. Charging time is nonlinear and modeled as a piecewise linear function.
5. Charging stations are available anytime, indicating continuous accessibility throughout the travel.

### Parameters

- `B`: Total energy capacity of the EV.
- `r_0`: Initial energy consumption, which should be sufficient to reach the first charging station while maintaining the SOC above `B_min` at the first station.
- `p_i`: Power rate at charging station `i`.
- `c_i`: Charging cost at station `i`, based on time-of-use (TOU) prices.
- `e_i`: Energy consumption required to reach charging station `i` and the destination.

The charging time follows a concave structure, increasing as the battery nears full charge, typically after reaching 80% of the battery capacity.

## Model Contributions

- **Novel MILP model:** Determines optimal charging strategies considering travel time and cost.
- **Single-stage decision-making:** Provides insights into when, where, and how much charge is needed.
- **Comprehensive modeling:** Accounts for multiple refueling stops, partial charging, nonlinear charging times, and varying power rates.

## Installation

To install and run this project, you need to clone the repository and ensure you have the necessary dependencies.

```bash
git clone <repository-url>
cd <repository-directory>
```

Install the required Python packages:

```bash
pip install -r requirements.txt
```

## Usage

To run the MILP model, execute the `main.py` script:

```bash
python main.py
```

Ensure that the problem data is properly configured in the script or provided as input.

## Example

An example problem setup is provided in `CSPM.py`. This script contains sample data for running the model. Modify the parameters as needed to test different scenarios.

## Requirements

To run the model, you need the following Python packages:

- **openpyxl**: For reading and writing Excel files, specifically for loading the problem data from an Excel file.
- **docplex**: IBM Decision Optimization CPLEX Modeling for Python, used for formulating and solving the MILP model.
- **collections**: To handle the data structure for problem elements.
- **namedtuple**: To create efficient and lightweight data structures.

These packages can be installed using the following command:

```bash
pip install gc openpyxl docplex collections
```

## References

1. Yılmaz & Yagmahan (2024), "Optimization of the Electric Vehicle Charging Strategy Problem for Sustainable Intercity Travels with Multiple Refueling Stops", Sustainable Energy, Grids and Networks, 


## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
