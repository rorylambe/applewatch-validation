# Apple Watch Series 10 VO₂max Validation

Analysis code and data for a laboratory validation of **VO₂max estimates from Apple Watch Series 10** against indirect calorimetry, the criterion method.

| Outcome | Index device | Criterion |
|---|---|---|
| VO₂max (mL·kg⁻¹·min⁻¹) | Apple Watch Series 10 estimate (watchOS 11), generated in free-living conditions | COSMED Quark CPET indirect calorimetry (maximal treadmill test, modified Åstrand protocol) |

The study enrolled **40 healthy adults** (20 women, 20 men; `participant01`–`participant40`). Data were collected in Dublin, Ireland, between October 2024 and January 2025.

---

## Folder contents

```
series10/
├── README.md
├── BlandAltman_mae_mape.ipynb   # Agreement analysis for VO₂max
└── vo2max_data.csv              # Participant-level criterion and device values
```

---

## Data dictionary

`vo2max_data.csv` contains one row per participant.

| Column | Description |
|---|---|
| `participant_id` | Pseudonymised participant ID (`participant01`–`participant40`) |
| `sex` | Participant sex (`m`/`f`) |
| `applewatch_model` | Apple Watch model worn (`S10` = Series 10) |
| `vo2max_attained` | Whether true VO₂max was attained during criterion testing (`Yes`/`No`) |
| `include_vo2` | Inclusion flag for the VO₂max analysis (`yes`/`no`) |
| `cosmed_vo2max` | Criterion VO₂max from COSMED (mL·kg⁻¹·min⁻¹) |
| `applewatch_vo2max` | Apple Watch Series 10 VO₂max estimate (mL·kg⁻¹·min⁻¹) |
| `cosmed_crf_percentile_friend` | Cardiorespiratory fitness percentile from the criterion VO₂max, using FRIEND registry reference standards |

Where no Apple Watch estimate was available, `applewatch_vo2max` holds a text code instead of a number (`no_prediction_obtained` or `not_obtained_within_timeframe`). These are coerced to `NaN` during processing.

### Inclusion

Participants were included in the analysis (`include_vo2 == 'yes'`) if they:

1. attained VO₂max during the criterion test, meeting at least 2 of the following: heart rate within ±10 bpm of age-predicted maximum (220 − age), respiratory exchange ratio ≥ 1.15, rating of perceived exertion ≥ 17, or a VO₂ plateau (< 150 mL/min increase with increasing work rate), **and**
2. generated an Apple Watch VO₂max estimate within 1 week of criterion testing.

Five of the 40 participants were excluded: 1 did not attain VO₂max, and 4 did not generate an Apple Watch estimate within the timeframe. **35 participants** are included in the final analysis.

---

## Analysis

The notebook `BlandAltman_mae_mape.ipynb` computes agreement between the Apple Watch Series 10 VO₂max estimate (index) and COSMED VO₂max (criterion).

**1. Data processing**
- Coerce VO₂max columns to numeric
- Drop participants with missing criterion or index values
- Restrict to participants meeting inclusion criteria (`include_vo2 == 'yes'`, n = 35)

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

The FRIEND percentile classification and the exploratory subgroup analyses (by sex and fitness level) reported in the paper are not included in this notebook.

---

## Results snapshot

From the saved notebook output (n = 35):

| Statistic | Value |
|---|---|
| Mean COSMED | 51.51 mL·kg⁻¹·min⁻¹ (SD 5.88) |
| Mean Apple Watch Series 10 | 45.26 mL·kg⁻¹·min⁻¹ (SD 7.64) |
| Bias (Apple Watch − COSMED) | −6.25 mL·kg⁻¹·min⁻¹ (95% CI −8.28 to −4.23) |
| SD of differences | 6.11 mL·kg⁻¹·min⁻¹ |
| Lower LoA | −18.23 mL·kg⁻¹·min⁻¹ (95% CI −21.73 to −14.74) |
| Upper LoA | 5.73 mL·kg⁻¹·min⁻¹ (95% CI 2.24 to 9.23) |
| MAE | 6.79 mL·kg⁻¹·min⁻¹ (95% CI 4.91 to 8.68) |
| MAPE | 13.2% (95% CI 9.6 to 16.8) |

Apple Watch Series 10 underestimated VO₂max on average, and the wide limits of agreement show substantial error for individual estimates.

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
cd applewatch-validation/series10
pip install pandas numpy scipy matplotlib plotly scikit-learn jupyter
jupyter notebook BlandAltman_mae_mape.ipynb
```

The notebook reads `vo2max_data.csv` from the working directory, so launch it from inside this folder.

---

## Reusing the workflow for other devices

Sections 2–4 are device-agnostic: once `criterion` and `index` are defined as paired pandas Series, the Bland-Altman, MAPE and MAE code runs unchanged. Only the data-loading and cleaning steps in Section 1 need adapting to a new dataset.

---

## Corresponding author

**Rory Lambe**, University College Dublin

## Citation

If you use this code or data, please cite the associated manuscript:
Lambe, Rory, et al. "Accuracy of VO2max Estimates From Apple Watch Series 10." Mayo Clinic Proceedings: Digital Health 4.2 (2026): 100357. https://doi.org/10.1016/j.mcpdig.2026.100357

## Licence

Code is released under the [MIT License](../LICENSE). Data (`vo2max_data.csv`) are released under [CC BY 4.0](../LICENSE-DATA).
