# 77 GHz SIW-Fed 4-Element ADAS Radar Antenna

A CST Studio Suite 2024 project investigating a compact Substrate Integrated Waveguide (SIW)-fed four-element patch antenna array for automotive Advanced Driver Assistance Systems (ADAS) radar and beam-steering applications.

## Project Overview

This project focuses on the electromagnetic design and simulation of a millimeter-wave antenna array operating at **77 GHz**, a frequency widely used in automotive radar systems.

The design combines an SIW feeding structure, a microstrip feed network, four patch radiators, and a single-stub phase network. The goal is to investigate impedance matching, radiation characteristics, and controlled beam steering through electromagnetic simulation.

## Design Architecture

- **Operating frequency:** 77 GHz
- **Antenna configuration:** Four-element patch array
- **Feeding structure:** Substrate Integrated Waveguide (SIW)
- **Feed network:** SIW-to-microstrip transitions and microstrip routing
- **Phase network:** Single-stub phase network
- **Simulation software:** CST Studio Suite 2024
- **Automation:** VBA geometry-generation macro
- **Intended application:** Automotive ADAS radar and beam-steering studies

## Initial Design Parameters

| Parameter | Value |
|---|---:|
| Target frequency | 77 GHz |
| Simulation sweep | 76–81 GHz |
| Substrate | RT/duroid 5880 |
| Relative permittivity | 2.2 |
| Substrate thickness | 0.127 mm |
| Copper thickness | 0.017 mm |
| Via diameter | 0.3 mm |
| Via pitch | 0.5 mm |
| Number of radiating elements | 4 |

*Note: These are the current design assumptions. Material properties and electromagnetic performance require appropriate verification.*

## Simulation Objectives

1. Develop the SIW-fed four-element antenna geometry.
2. Validate the input-port excitation and solver configuration.
3. Evaluate input impedance and reflection coefficient \(S_{11}\).
4. Investigate impedance matching near 77 GHz.
5. Analyze far-field radiation patterns, directivity, gain, and efficiency.
6. Investigate broadside operation and intended beam-steering states.
7. Assess mesh convergence and boundary-condition sensitivity.

## Current Validation Status

The CST Time Domain solver has completed successfully in the latest recorded run, and the corresponding S-parameter and VSWR results were generated.

However, the antenna is **still under development and validation**.

The current S₁₁ response has a minimum of approximately −14 dB near 80.6 GHz, while the value near 77 GHz is approximately −0.7 dB. These values are visual estimates from the current simulation plot.

The displayed far-field results also require further validation at the intended 77 GHz operating frequency. The feed-network phase balance, radiation pattern, and beam-steering behavior remain under investigation.

No final antenna performance claims are made at this stage.

## Repository Structure

The intended repository organization is:

```text
ADAS-Radar-Antenna/
├── README.md
├── CST/
│   ├── antenna_model.cst
│   └── build_antenna.bas
├── Results/
│   ├── S11_77GHz.png
│   └── Farfield_77GHz.png
└── Documentation/
    └── Design_Notes.md
```

*The structure above is a proposed organization. Actual filenames and folders may differ.*

## Tools and Technologies

- CST Studio Suite 2024
- VBA macro-based geometry generation
- SIW and microstrip transmission-line design
- Millimeter-wave antenna simulation
- S-parameter and far-field analysis

## Future Work

- Validate the input-port mode and boundary configuration.
- Investigate the matching shift relative to 77 GHz.
- Evaluate patch dimensions and feed-network discontinuities.
- Verify phase balance across the four radiating elements.
- Validate broadside radiation and intended steering angles.
- Perform mesh-convergence and simulation-sensitivity studies.
- Document final validated electromagnetic performance.

## Disclaimer

This repository documents an ongoing engineering simulation project. Numerical results are preliminary and should not be interpreted as validated hardware performance. Further simulation, verification, and potentially experimental measurements are required before drawing final conclusions.

## Author

Developed as an engineering project focused on millimeter-wave antenna design, electromagnetic simulation, and automotive radar applications.
