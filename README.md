# Insurance Claims Analysis & VBA Automation

## Project Overview

This project analyzes insurance claim data using Microsoft Excel and VBA.

The project focuses on building an end-to-end workflow that starts with raw insurance claims data and progresses through data cleaning, data-quality validation, claims analysis, fraud analysis and dashboard visualization.

A VBA-based workflow was developed to automate repetitive data preparation and analysis tasks, allowing the cleaned dataset and analytical outputs to be regenerated efficiently.

---

# 📊 Dataset
The dataset contains claim-level insurance observations. The main variables include:

 `claim_id`, `customer_id`, `incident_cause`, `claim_date`, `claim_area`, `police_report`, `claim_type`, `claim_amount`, `total_policy_claims`, `fraudulent` 

---

# 🔄 Data Preparation & Cleaning
The project follows a structured workflow to preserve the original data while creating a separate cleaned dataset.
### Project workflow

Raw Data  ->  VBA Data Cleaning  ->  Data Quality Validation  ->  Derived Variables  ->  Claims Analysis  >  Fraud & Customer Analysis  ->  Dashboard

---

## 1. Raw Data
The original dataset is kept separately in the `Raw_Data` worksheet. The raw dataset is not directly modified during the analysis process.
Instead, a separate `Clean_Data` worksheet is created for processing. This allows the original data to be preserved and provides a clear distinction between raw and processed data.

### Raw data structure

![image alt](https://github.com/kristineng/VBA-Automation-Insurance-Claims-Analysis/blob/4dd804cab047efcf9263968b37d476b7f1af48fc/images/raw_data.png)

---

## 2. VBA Data Cleaning
A VBA macro was developed to automate the cleaning process.
The cleaning workflow includes:

- Copying the raw dataset into a separate cleaned-data worksheet
- Cleaning monetary claim amounts
- Converting claim amounts into numeric values
- Standardizing claim date values
- Checking missing claim amounts
- Checking missing policy claim counts
- Identifying duplicate claim IDs
- Identifying invalid claim amounts
- Identifying invalid dates

For example, claim amounts originally stored with currency symbols such as: `$2,980` are converted into numeric values so that they can be used correctly in calculations and analysis. The formatted value is then displayed as: `$2,980.00`

---

## 3. Data Quality Validation
Rather than automatically deleting problematic records, the project records data-quality issues and preserves the observations for further review. A `data_quality_flag` variable was created to identify records requiring attention.
A separate `Data_Quality` worksheet summarizes the number of records affected by each type of issue.

![image alt](https://github.com/kristineng/VBA-Automation-Insurance-Claims-Analysis/blob/4dd804cab047efcf9263968b37d476b7f1af48fc/images/data_quality.png)

This approach avoids treating missing financial information as zero and makes the data-quality process transparent.

---

## 4. Derived Variables
A `claim_severity` variable was created based on claim amount.
The categories are:

- Low
- Medium
- High
- Very High

This variable allows claim amounts to be analyzed using categorical severity levels in addition to the original numerical claim amount.

---

# 📈 Claims Analysis
After cleaning and validating the data, the project analyzes claim characteristics and claim amounts.
The analysis includes:

- Total number of claims
- Total claim amount
- Average claim amount
- Median claim amount
- Maximum claim amount
- Minimum claim amount

---

## Claims by Claim Type
Claims are grouped by claim type to compare both claim frequency and total claim amount.

![image alt](https://github.com/kristineng/VBA-Automation-Insurance-Claims-Analysis/blob/4dd804cab047efcf9263968b37d476b7f1af48fc/images/claim_type.png)

---

## Claims by Incident Cause
Claims are also analyzed according to the underlying incident cause. This allows differences in claim frequency and claim amounts across incident causes to be examined.

![image alt](https://github.com/kristineng/VBA-Automation-Insurance-Claims-Analysis/blob/4dd804cab047efcf9263968b37d476b7f1af48fc/images/incident_cause.png)

---

## Claims by Area
The dataset is grouped by claim area to compare claim frequency and financial exposure across different areas.

![image alt](https://github.com/kristineng/VBA-Automation-Insurance-Claims-Analysis/blob/4dd804cab047efcf9263968b37d476b7f1af48fc/images/area.png)

---

## Claims by Police Report

Claims are analyzed according to the police report indicator. Both the number of claims and associated claim amounts are compared.

![image alt](https://github.com/kristineng/VBA-Automation-Insurance-Claims-Analysis/blob/4dd804cab047efcf9263968b37d476b7f1af48fc/images/police_report.png)

---

## Claims by Severity
The derived claim severity variable is used to examine the distribution of claims across severity categories.

![image alt](https://github.com/kristineng/VBA-Automation-Insurance-Claims-Analysis/blob/4dd804cab047efcf9263968b37d476b7f1af48fc/images/severity.png)

---

# 👤 Customer & Policy Analysis
The project also examines claim behavior at the customer level.
The analysis includes:
- Total number of customers
- Average claims per customer
- Number of customers with multiple claims

This provides an additional perspective beyond claim-level analysis by examining how claims are distributed across customers.

---

# 📅 Time Analysis
Claim dates are used to analyze changes in claim activity over time.
The analysis includes:
- Claims by year
- Total claim amount by year
- Average claim amount by year

![image alt](https://github.com/kristineng/VBA-Automation-Insurance-Claims-Analysis/blob/053dfaf983068339163ea0ef7734957ffcb2a05e/images/time_analysis.png)

This allows changes in claim frequency and financial exposure to be compared across different periods.

---

# 🚨 Fraud Analysis
Fraudulent and non-fraudulent claims are analyzed separately.
The fraud analysis includes:
- Number of fraudulent claims
- Fraud rate
- Total fraudulent claim amount
- Average fraudulent claim amount
- Comparison of fraudulent and non-fraudulent claims

The comparison considers both claim frequency and claim amount.

![image alt](https://github.com/kristineng/VBA-Automation-Insurance-Claims-Analysis/blob/053dfaf983068339163ea0ef7734957ffcb2a05e/images/fraud_analysis.png)

---

# 📊 Dashboard
The final dashboard brings together the major findings from the analysis.
The dashboard includes key performance indicators and visualizations covering:

- Overall claim activity
- Claim amounts
- Claim characteristics
- Claim severity
- Fraud
- Time trends
- Other portfolio-level measures

![image alt](https://github.com/kristineng/VBA-Automation-Insurance-Claims-Analysis/blob/053dfaf983068339163ea0ef7734957ffcb2a05e/images/dashboard_vba.png)

The dashboard is designed to provide a high-level view of the insurance claims portfolio while allowing the underlying analysis to be traced back to the `Analysis` worksheet.

---

# ⚙️ VBA Automation
VBA was used to automate the data-processing and analysis workflow.
The main automation components include:
### `CleanClaimsData`
- Creating the cleaned dataset
- Cleaning claim amounts
- Validating dates
- Checking missing values
- Identifying duplicate claim IDs
- Creating claim severity
- Creating row-level data-quality flags
- Generating the data-quality summary
### `ClaimsAnalysis`
Responsible for generating the main claims analysis, including:
- Portfolio-level metrics
- Claims by claim type
- Claims by incident cause
- Fraud overview
- Claims by area
- Claims by police report
- Claims by severity
- Customer/policy analysis
- Time analysis
- Fraud analysis

The VBA workflow reduces the need to manually repeat the same calculations when the underlying data changes.

---

Thank you!
