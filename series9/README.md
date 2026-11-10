# Apple Watch VO₂max Validation (Series 9 and other models)

Analysis code and data for a laboratory validation of **VO₂max estimates from Apple Watch** against indirect calorimetry, the criterion method.

| Outcome | Index device | Criterion |
|---|---|---|
| VO₂max (mL·kg⁻¹·min⁻¹) | Apple Watch estimate (watchOS 10 or later), generated in free-living conditions | COSMED Quark CPET indirect calorimetry (maximal treadmill test, modified Åstrand protocol) |

The study enrolled **30 healthy adults** (15 women, 15 men). Data were collected in Dublin, Ireland, between December 2023 and July 2024.

Participants who did not already own an Apple Watch were given a Series 9 or Ultra 2 for 5–10 days. Others used their own watch, so several models are included:

| Model | Participants |
|---|---|
| Series 9 | 17 |
| Ultra 2 | 6 |
| Series 8 | 2 |
| SE 2 | 2 |
| Series 5, 6 and 7 | 1 each |

---

## Folder contents

```
series9/
├── README.md
├── BlandAltman_mae_mape.ipynb   # Agreement analysis for VO₂max
└── vo2max_data.csv              # Participant-level criterion and device values
```

---

## Data dictionary

`vo2max_data.csv` contains one row per participant.

| Column | Description |
|---|---|
| `participant_id` | Pseudonymised participant ID |
| `applewatch_model` | Apple Watch model worn (e.g. `Series 9`, `Ultra 2`, `SE 2`) |
| `VO2max attained?` | Whether true VO₂max was attained during criterion testing (`Y`/`N`) |
| `include_vo2` | Inclusion flag for the VO₂max analysis (`yes`/`no`) |
| `cosmed_vo2max` | Criterion VO₂max from COSMED (mL·kg⁻¹·min⁻¹) |
| `applewatch_vo2max` | Apple Watch VO₂max estimate (mL·kg⁻¹·min⁻¹) |

Rows with missing values are removed during processing.

### Inclusion

Participants were included in the analysis (`include_vo2 == 'yes'`) if they attained VO₂max during the criterion test, meeting at least 2 of the following: heart rate within ±10 bpm of age-predicted maximum (220 − age), respiratory exchange ratio ≥ 1.15, rating of perceived exertion ≥ 17, or a VO₂ plateau (< 150 mL/min increase with increasing work rate).

Two of the 30 participants did not attain VO₂max and were excluded, so **28 participants** are included in the final analysis.

---

## Analysis

The notebook `BlandAltman_mae_mape.ipynb` computes agreement between the Apple Watch VO₂max estimate (index) and COSMED VO₂max (criterion).

**1. Data processing**
- Coerce VO₂max columns to numeric
- Drop participants with missing criterion or index values
- Restrict to participants meeting inclusion criteria (`include_vo2 == 'yes'`, n = 28)

**2. Bland-Altman limits of agreement**
- Bias computed as **index − criterion** (Apple Watch − COSMED)
- 95% limits of agreement (bias ± 1.96 SD)
- 95% CI for the bias and for each limit of agreement, using SE(LoA) = SD × √(1/n + 1.96² / (2(n − 1)))
- Bland-Altman plot

**3. Mean absolute percentage error (MAPE)**
- Function-based and step-by-step manual calculation
- 95% CI using the *t*-distribution (small-sample)

**4. Mean absolute error (MAE)**
- Manual calculation, cross-checked against `sklearn.metrics.mean_absolute_error`
- 95% CI using the *t*-distribution

---

## Results snapshot

From the saved notebook output (n = 28):

| Statistic | Value |
|---|---|
| Mean COSMED | 49.35 mL·kg⁻¹·min⁻¹ (SD 7.79) |
| Mean Apple Watch | 43.27 mL·kg⁻¹·min⁻¹ (SD 6.34) |
| Bias (Apple Watch − COSMED) | −6.07 mL·kg⁻¹·min⁻¹ (95% CI −8.38 to −3.77) |
| SD of differences | 6.22 mL·kg⁻¹·min⁻¹ |
| Lower LoA | −18.26 mL·kg⁻¹·min⁻¹ (95% CI −22.24 to −14.28) |
| Upper LoA | 6.11 mL·kg⁻¹·min⁻¹ (95% CI 2.13 to 10.09) |
| MAE | 6.92 mL·kg⁻¹·min⁻¹ (95% CI 4.89 to 8.94) |
| MAPE | 13.31% (95% CI 10.01 to 16.61) |

Apple Watch underestimated VO₂max on average, and the wide limits of agreement show substantial error for individual estimates.

**Note on sign convention:** the published paper reports the bias as COSMED − Apple Watch (mean difference 6.07; LoA −6.11 to 18.26). This notebook uses Apple Watch − COSMED, matching the Series 10 analysis, so the signs are reversed. The magnitudes are identical.

---

## Getting started

**Requirements:** Python 3.13 (developed on 3.13.7) and the following packages:

```
pandas
numpy
scipy
matplotlib
plotly
scikit-learn
jupyter
```

**Run:**

```bash
git clone https://github.com/rorylambe/applewatch-validation.git
cd applewatch-validation/series9
pip install pandas numpy scipy matplotlib plotly scikit-learn jupyter
jupyter notebook BlandAltman_mae_mape.ipynb
```

The notebook reads `vo2max_data.csv` from the working directory, so launch it from inside this folder.

---

## Corresponding author

**Rory Lambe**, University College Dublin

## Citation

If you use this code or data, please cite the associated manuscript:
Lambe, Rory, et al. "Investigating the accuracy of Apple Watch VO2 max measurements: A validation study." PLOS ONE 20.5 (2025): e0323741. https://doi.org/10.1371/journal.pone.0323741

## Licence

Code is released under the [MIT License](../LICENSE). Data (`vo2max_data.csv`) are released under [CC BY 4.0](../LICENSE-DATA).
