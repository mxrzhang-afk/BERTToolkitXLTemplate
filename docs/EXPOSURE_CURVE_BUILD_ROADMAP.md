# Exposure Curve Build Roadmap

Branch: `Exposure_Curve_Buid`

## Objective

Define and implement the `build` action for the `<<Profile>>` tab. The build
action converts exposure profile inputs into XLsimulation-compatible CDF loss
cause input tables.

## V1 Scope

`profile_build` reads `Input_Profile` and writes one CSV per loss cause name in
the exposure rating parameter table.

Output folder:

```text
<workbook folder>/_BERTToolkitTemp/xl_pricing_tool/profile/build/
```

Output file naming:

```text
<LossCause>_cdf.csv
```

Examples:

```text
P_Expo_PptyMB_cdf.csv
C_Expo_PptyMB_cdf.csv
P_Expo_CEAR_cdf.csv
C_Expo_CEAR_cdf.csv
```

No manifest is planned for V1.

## Input Contract

Risk profile rows:

```text
Input_Profile!B12:H...
AsAt | UY | LOB | Nrisk | SI | Premium | Incurred
```

Build controls:

```text
Input_Profile!M4 = Minimum Loss / threshold
Input_Profile!M5 = Loss Cap
Input_Profile!M6 = Percentile Steps
```

Exposure rating parameters:

```text
Input_Profile!K10:U32
```

Prior rows are read from `M13:U20`; current rows are read from `M24:U31`.

Used columns:

```text
M = LOB
N = LossCause
O = ELR
P = SubjectPrem / projected premium
Q = ActualPrem
R = Curve name
S = CurrencyAdj
T = b / curve parameter 1
U = g / curve parameter 2
```

Rows with blank LOB or blank LossCause are skipped.

## Curve Families

Property and CEAR risk use MBBEFD curves from `MBCurves`:

```text
A = Curve
B = b
C = g
```

Casualty lines use increased limit factor curves from `ILFCurves`:

```text
A = Curve
B = LOB
C = Default
D = State
E = Table
F = GCIdx
G:R = mu_1 ... mu_12
S:AD = w_1 ... w_12
```

V1 identifies the curve family by curve name:

- present only in `MBCurves` => `MBBEFD`
- present only in `ILFCurves` => `ILF`
- present in both or neither => validation error

## XLsimulation CSV Layout

Each output CSV must match the XLsimulation `CDF` loss cause source block:

```text
Parameter,Value
Mean Frequency,<derived>
Frequency Type,Poisson
Interpolation,Linear
Loss Cap,<Input_Profile!M5>
Minimum Loss,<Input_Profile!M4>
Percentile,Loss Severity
0,<severity>
...
1,<severity>
```

If `M6 = n`, output `n + 1` percentile rows including both `0` and `1`,
plus the `Parameter,Value` header row and five parameter rows.

## Frequency Derivation

For each loss cause:

```text
ExpectedLoss = SubjectPrem * ELR
MeanFrequency = ExpectedLoss / ConditionalMeanSeverity
```

`SubjectPrem` from column `P` is the premium basis. `ActualPrem` from column
`Q` is retained as context but is not used for V1 frequency derivation.

## Severity Derivation

MBBEFD rows are applied to profile rows for the matching `AsAt` and `LOB`.

For each profile row:

```text
AvgSI = SI / Nrisk
DamageRatio ~ MBBEFD(g, b)
Loss = AvgSI * DamageRatio
```

The aggregate severity CDF is the `Nrisk`-weighted mixture of row-level damage
ratio distributions.

ILF rows use the mixed-exponential severity model from `ILFCurves`:

```text
F(x) = sum_i w_i * (1 - exp(-x / mu_i))
```

Weights are normalized after dropping nonpositive weights and nonpositive
`mu` values.

Both curve families are inverted numerically onto the requested percentile
grid conditional on losses falling between `Minimum Loss` and `Loss Cap`.
Losses below `Minimum Loss` are ignored rather than floored; the generated
frequency is therefore an exceedance frequency and the generated severity CDF is
conditional on exceeding `Minimum Loss`.

## Out Of Scope For V1

- No workbook writeback from `profile_build`.
- No manifest file.
- No automatic loading into `Sim_Variations`.
- No use of `ActualPrem` beyond validation/context.
- No special handling for `CurrencyAdj` unless later specified.
