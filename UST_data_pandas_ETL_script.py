# ================================================================================================================
#               ETL FOR UNIVERSITY OF ST. THOMAS UNDERGRAD ADMISSIONS/RETENTION/GRADUATION DATA
# ================================================================================================================


# IMPORT
import pandas as pd


# DEFINE EXCEL FILE NAMES
read_data_file_name = '20952596_admission_grad_retention_TB.xlsx'
write_data_file_name = 'ust_admission_grad_retention_data.xlsx'


# CREATE "FIND_VALUE" FUNCTION
def find_value(df, row_filter1, row_criteria1, row_filter2, row_criteria2, col_name):
    """Returns a cell value that matches 2 row criteria, given a column name."""
    cell_value = df.loc[((df[row_filter1] == row_criteria1) & (df[row_filter2] == row_criteria2)), col_name].values[0]
    return cell_value


# DEFINE VALUES LISTS FOR GENDER AND RACE CATEGORIES
gender_list = ['Men', 'Women']
race_list = ['Nonresident aliens', 'American Indian or Alaskan Native', 'Race and ethnicity unknown',
             'Hispanics of any race', 'Asian', 'Black or African American', 'Native Hawaiian or Other Pacific Islander',
             'White', 'Two or more races']


# DEFINE HEADER VALUES FOR DATAFRAMES
retention_by_group_headers = ['index_one', 'index_two', 'cohort_name', 'cohort_original',
                                 'cohort_modified', 'ret_2nd_yr', 'ret_3rd_yr', 'ret_4th_yr', 'ret_5th_yr',
                                 'ret_6th_yr', 'ret_7th_yr', 'ret_8th_yr', 'ret_9th_yr', 'ret_10th_yr', 'ret_11th_yr']

grad_by_group_headers = ['index_one', 'index_two', 'cohort_name', 'cohort_original',
                                 'cohort_modified', 'grad_in_1yr', 'grad_in_2yr', 'grad_in_3yr',
                                 'grad_in_4yr', 'grad_in_5yr', 'grad_in_6yr', 'grad_in_7yr',
                                 'grad_in_8yr', 'grad_in_9yr', 'grad_in_10yr']

retention_total_headers = ['cohort_name', 'cohort_original', 'cohort_modified', 'ret_2nd_yr',
                                   'ret_3rd_yr', 'ret_4th_yr', 'ret_5th_yr', 'ret_6th_yr', 'ret_7th_yr',
                                   'ret_8th_yr', 'ret_9th_yr', 'ret_10th_yr', 'ret_11th_yr']

grad_total_headers = ['cohort_name', 'cohort_original', 'cohort_modified', 'grad_in_1yr',
                         'grad_in_2yr',  'grad_in_3yr', 'grad_in_4yr', 'grad_in_5yr', 'grad_in_6yr',
                         'grad_in_7yr', 'grad_in_8yr', 'grad_in_9yr', 'grad_in_10yr']

admissions_headers = ['status', 'gender', 'fall_2017', 'fall_2018', 'fall_2019', 'fall_2020', 'fall_2021']

rates_headers = ['rate', 'gender', 'fall_2017', 'fall_2018', 'fall_2019', 'fall_2020', 'fall_2021']


# ================================== RETENTION BY GENDER AND RACE ==================================================



# IMPORT "RETENTION BY GENDER AND RACE" SHEET FROM EXCEL AS DATAFRAME
df_retention_by_group_temp = pd.read_excel(read_data_file_name,
                                    sheet_name='4yr-FTFY_retention_gender_race',
                                      header=[4], usecols='A:O')



# RENAME COLUMNS IN "RETENTION_BY_GROUP" DATAFRAME
df_retention_by_group_temp.columns = retention_by_group_headers



# FORWARD FILL COLUMNS FOR FIRST TWO COLUMNS WHICH CONTAIN MERGED ROWS
cols = ['index_one', 'index_two']
df_retention_by_group_temp.loc[:,cols] = df_retention_by_group_temp.loc[:,cols].ffill()



# GET RID OF BLANK ROWS AND EXTRA HEADER ROWS IN DATAFRAME
df_retention_by_group = df_retention_by_group_temp.loc[df_retention_by_group_temp['cohort_name'].notnull()]



# CREATE DATAFRAME FOR RETENTION BY GENDER ONLY
df_retention_gender_only_subset = df_retention_by_group[((df_retention_by_group['index_one'] == 'Men')
                                                  | (df_retention_by_group['index_one'] == 'Women'))
                                                 & (df_retention_by_group['index_two'].isna())]

# remove "index_two" column from "gender only" dataframe, which is empty:
df_retention_gender_only_drop_col = df_retention_gender_only_subset.drop(columns='index_two')

# rename "index_one" column to "gender":
df_retention_gender_only_rename_col = df_retention_gender_only_drop_col.rename(columns={'index_one': 'gender'})

# rename "gender only" dataframe to final name:
df_retention_gender_only = df_retention_gender_only_rename_col

# print check:
# print(df_retention_gender_only)



# CREATE DATAFRAME FOR RETENTION BY BOTH GENDER AND RACE
df_retention_gender_race_subset = df_retention_by_group[((df_retention_by_group['index_one'] == 'Men')
                                                  | (df_retention_by_group['index_one'] == 'Women'))
                                                 & (df_retention_by_group['index_two'].notna())]

# rename "index_one" column to "gender" and "index_two" column to "race":
df_retention_gender_race_rename_col = df_retention_gender_race_subset.rename(columns={'index_one': 'gender', 'index_two': 'race'})

# rename "gender and race" dataframe to final name:
df_retention_gender_race = df_retention_gender_race_rename_col

# print check:
# print(df_retention_gender_race)



# CREATE DATAFRAME FOR RETENTION BY RACE ONLY
df_retention_race_only_subset = df_retention_by_group[df_retention_by_group['index_one'].isin(race_list)]

# remove "index_two" column of "race only" dataframe, which has incorrect values:
df_retention_race_only_drop_col = df_retention_race_only_subset.drop(columns='index_two')

# rename "index_one" column to "race":
df_retention_race_only_rename_col = df_retention_race_only_drop_col.rename(columns={'index_one': 'race'})

# rename "race only" dataframe to final name:
df_retention_race_only = df_retention_race_only_rename_col

# print check:
# print(df_retention_race_only)



# ================================== GRADUATION BY GENDER AND RACE ==================================================



# IMPORT "GRADUATION BY GENDER AND RACE" SHEET FROM EXCEL AS DATAFRAME
df_grad_by_group_temp = pd.read_excel(read_data_file_name,
                                    sheet_name='4yr-FTFY_grad_gender_race',
                                      header=[3], usecols='A:O')



# RENAME COLUMNS IN "GRAD_BY_GROUP" DATAFRAME
df_grad_by_group_temp.columns = grad_by_group_headers



# FORWARD FILL COLUMNS FOR FIRST TWO COLUMNS WHICH CONTAIN MERGED ROWS
cols = ['index_one', 'index_two']
df_grad_by_group_temp.loc[:,cols] = df_grad_by_group_temp.loc[:,cols].ffill()



# GET RID OF BLANK ROWS AND EXTRA HEADER ROWS IN DATAFRAME
df_grad_by_group = df_grad_by_group_temp.loc[df_grad_by_group_temp['cohort_name'].notnull()]



# CREATE DATAFRAME FOR GRADUATION BY GENDER ONLY
df_grad_gender_only_subset = df_grad_by_group[((df_grad_by_group['index_one'] == 'Men')
                                                  | (df_grad_by_group['index_one'] == 'Women'))
                                                 & (df_grad_by_group['index_two'].isna())]

# remove "index_two" column from "gender only" dataframe, which is empty:
df_grad_gender_only_drop_col = df_grad_gender_only_subset.drop(columns='index_two')

# rename "index_one" column to "gender":
df_grad_gender_only_rename_col = df_grad_gender_only_drop_col.rename(columns={'index_one': 'gender'})

# rename "gender only" dataframe to final name:
df_grad_gender_only = df_grad_gender_only_rename_col

# print check:
# print(df_grad_gender_only)



# CREATE DATAFRAME FOR GRADUATION BY BOTH GENDER AND RACE
df_grad_gender_race_subset = df_grad_by_group[((df_grad_by_group['index_one'] == 'Men')
                                                  | (df_grad_by_group['index_one'] == 'Women'))
                                                 & (df_grad_by_group['index_two'].notna())]

# rename "index_one" column to "gender" and "index_two" column to "race":
df_grad_gender_race_rename_col = df_grad_gender_race_subset.rename(columns={'index_one': 'gender', 'index_two': 'race'})

# rename "gender and race" dataframe to final name:
df_grad_gender_race = df_grad_gender_race_rename_col

# print check:
# print(df_grad_gender_race)



# CREATE DATAFRAME FOR GRADUATION BY RACE ONLY
df_grad_race_only_subset = df_grad_by_group[df_grad_by_group['index_one'].isin(race_list)]

# remove "index_two" column of "race only" dataframe, which has incorrect values:
df_grad_race_only_drop_col = df_grad_race_only_subset.drop(columns='index_two')

# rename "index_one" column to "race":
df_grad_race_only_rename_col = df_grad_race_only_drop_col.rename(columns={'index_one': 'race'})

# rename "race only" dataframe to final name:
df_grad_race_only = df_grad_race_only_rename_col

# print check:
# print(df_grad_race_only)



# ============================================ RETENTION TOTAL ======================================================



# IMPORT "RETENTION TOTAL" SHEET FROM EXCEL AS DATAFRAME
df_retention_total = pd.read_excel(read_data_file_name,
                                    sheet_name='4yr-FTFY_retention',
                                      header=[3], usecols='A:M')



# RENAME COLUMNS IN "RETENTION_TOTAL" DATAFRAME
df_retention_total.columns = retention_total_headers

# print check:
# print(df_retention_total)



# ============================================ GRADUATION TOTAL ======================================================



# IMPORT "GRADUATION TOTAL" SHEET FROM EXCEL AS DATAFRAME
df_grad_total = pd.read_excel(read_data_file_name,
                                    sheet_name='4yr-FTFY_graduation',
                                      header=[3], usecols='A:M')



# RENAME COLUMNS IN "RETENTION_TOTAL" DATAFRAME
df_grad_total.columns = grad_total_headers

# print check:
# print(df_grad_total)



# ========================================== ADMISSIONS ===============================================================



# IMPORT "ADMISSIONS" SHEET FROM EXCEL AS DATAFRAME
df_admissions_temp = pd.read_excel(read_data_file_name,
                                    sheet_name='4yr_admission_rate',
                                      header=[3], usecols='A:G')



# RENAME COLUMNS IN "ADMISSIONS" DATAFRAME
df_admissions_temp.columns = admissions_headers



# FORWARD FILL COLUMN FOR FIRST COLUMN WHICH CONTAINS MERGED ROWS
col = ['status']
df_admissions_temp.loc[:,col] = df_admissions_temp.loc[:,col].ffill()


#
# # RENAME "INDEX_ONE" COLUMN TO "STATUS" AND "INDEX_TWO" COLUMN TO "GENDER"
# df_admissions_rename_col = df_admissions_temp.rename(columns={'index_one': 'status', 'index_two': 'gender'})



# RENAME "APPLIED (COMPLETED APPS)" VALUE IN "STATUS" COLUMN TO "APPLIED"
df_admissions_rename_value = df_admissions_temp.replace({'Applied (Completed Apps)': 'Applied'})



# RENAME "ADMISSIONS" DATAFRAME TO FINAL NAME
df_admissions_total = df_admissions_rename_value

# print check:
# print(df_admissions_total)



# ========================================== ADMISSION RATE ===========================================================



# CREATE CALCULATED VALUES FOR "ADMISSION_RATE" DF (WHICH WILL BE CONCATENATED TO ORIGINAL ADMISSIONS DF LATER)


# ~~~~~~~~~~~~~~~ APPLIED COUNTS ~~~~~~~~~~~~~~~

# APPLIED COUNT FOR MEN BY YEAR

# men applied count in Fall 2017:
app_men_f17 = find_value(df_admissions_total, 'gender', 'Men', 'status', 'Applied', 'fall_2017')

# men applied count in Fall 2018:
app_men_f18 = find_value(df_admissions_total, 'gender', 'Men', 'status', 'Applied', 'fall_2018')

# men applied count in Fall 2019:
app_men_f19 = find_value(df_admissions_total, 'gender', 'Men', 'status', 'Applied', 'fall_2019')

# men applied count in Fall 2020:
app_men_f20 = find_value(df_admissions_total, 'gender', 'Men', 'status', 'Applied', 'fall_2020')

# men applied count in Fall 2021:
app_men_f21 = find_value(df_admissions_total, 'gender', 'Men', 'status', 'Applied', 'fall_2021')



# APPLIED COUNT FOR WOMEN BY YEAR

# women applied count in Fall 2017:
app_women_f17 = find_value(df_admissions_total, 'gender', 'Women', 'status', 'Applied', 'fall_2017')

# women applied count in Fall 2018:
app_women_f18 = find_value(df_admissions_total, 'gender', 'Women', 'status', 'Applied', 'fall_2018')

# women applied count in Fall 2019:
app_women_f19 = find_value(df_admissions_total, 'gender', 'Women', 'status', 'Applied', 'fall_2019')

# women applied count in Fall 2020:
app_women_f20 = find_value(df_admissions_total, 'gender', 'Women', 'status', 'Applied', 'fall_2020')

# women applied count in Fall 2021:
app_women_f21 = find_value(df_admissions_total, 'gender', 'Women', 'status', 'Applied', 'fall_2021')



# APPLIED COUNT FOR "NOT REPORTED" BY YEAR

# "not reported" applied count in Fall 2017:
app_notreported_f17 = find_value(df_admissions_total, 'gender', 'Not Reported', 'status', 'Applied', 'fall_2017')

# "not reported" applied count in Fall 2018:
app_notreported_f18 = find_value(df_admissions_total, 'gender', 'Not Reported', 'status', 'Applied', 'fall_2018')

# "not reported" applied count in Fall 2019:
app_notreported_f19 = find_value(df_admissions_total, 'gender', 'Not Reported', 'status', 'Applied', 'fall_2019')

# "not reported" applied count in Fall 2020:
app_notreported_f20 = find_value(df_admissions_total, 'gender', 'Not Reported', 'status', 'Applied', 'fall_2020')

# "not reported" applied count in Fall 2021:
app_notreported_f21 = find_value(df_admissions_total, 'gender', 'Not Reported', 'status', 'Applied', 'fall_2021')



# APPLIED COUNT TOTAL BY YEAR

# total applied count in Fall 2017:
app_total_f17 = find_value(df_admissions_total, 'gender', 'Total', 'status', 'Applied', 'fall_2017')

# total applied count in Fall 2018:
app_total_f18 = find_value(df_admissions_total, 'gender', 'Total', 'status', 'Applied', 'fall_2018')

# total applied count in Fall 2019:
app_total_f19 = find_value(df_admissions_total, 'gender', 'Total', 'status', 'Applied', 'fall_2019')

# total applied count in Fall 2020:
app_total_f20 = find_value(df_admissions_total, 'gender', 'Total', 'status', 'Applied', 'fall_2020')

# total applied count in Fall 2021:
app_total_f21 = find_value(df_admissions_total, 'gender', 'Total', 'status', 'Applied', 'fall_2021')


# ~~~~~~~~~~~~~~~ ADMITTED COUNTS ~~~~~~~~~~~~~~~

# ADMITTED COUNT FOR MEN BY YEAR

# men admitted count in Fall 2017:
admt_men_f17 = find_value(df_admissions_total, 'gender', 'Men', 'status', 'Admitted', 'fall_2017')

# men admitted count in Fall 2018:
admt_men_f18 = find_value(df_admissions_total, 'gender', 'Men', 'status', 'Admitted', 'fall_2018')

# men admitted count in Fall 2019:
admt_men_f19 = find_value(df_admissions_total, 'gender', 'Men', 'status', 'Admitted', 'fall_2019')

# men admitted count in Fall 2020:
admt_men_f20 = find_value(df_admissions_total, 'gender', 'Men', 'status', 'Admitted', 'fall_2020')

# men admitted count in Fall 2021:
admt_men_f21 = find_value(df_admissions_total, 'gender', 'Men', 'status', 'Admitted', 'fall_2021')



# ADMITTED COUNT FOR WOMEN BY YEAR

# women admitted count in Fall 2017:
admt_women_f17 = find_value(df_admissions_total, 'gender', 'Women', 'status', 'Admitted', 'fall_2017')

# women admitted count in Fall 2018:
admt_women_f18 = find_value(df_admissions_total, 'gender', 'Women', 'status', 'Admitted', 'fall_2018')

# women admitted count in Fall 2019:
admt_women_f19 = find_value(df_admissions_total, 'gender', 'Women', 'status', 'Admitted', 'fall_2019')

# women admitted count in Fall 2020:
admt_women_f20 = find_value(df_admissions_total, 'gender', 'Women', 'status', 'Admitted', 'fall_2020')

# women admitted count in Fall 2021:
admt_women_f21 = find_value(df_admissions_total, 'gender', 'Women', 'status', 'Admitted', 'fall_2021')



# ADMITTED COUNT FOR "NOT REPORTED" BY YEAR

# "not reported" admitted count in Fall 2017:
admt_notreported_f17 = find_value(df_admissions_total, 'gender', 'Not Reported', 'status', 'Admitted', 'fall_2017')

# "not reported" admitted count in Fall 2018:
admt_notreported_f18 = find_value(df_admissions_total, 'gender', 'Not Reported', 'status', 'Admitted', 'fall_2018')

# "not reported" admitted count in Fall 2019:
admt_notreported_f19 = find_value(df_admissions_total, 'gender', 'Not Reported', 'status', 'Admitted', 'fall_2019')

# "not reported" admitted count in Fall 2020:
admt_notreported_f20 = find_value(df_admissions_total, 'gender', 'Not Reported', 'status', 'Admitted', 'fall_2020')

# "not reported" admitted count in Fall 2021:
admt_notreported_f21 = find_value(df_admissions_total, 'gender', 'Not Reported', 'status', 'Admitted', 'fall_2021')



# ADMITTED COUNT TOTAL BY YEAR

# total admitted count in Fall 2017:
admt_total_f17 = find_value(df_admissions_total, 'gender', 'Total', 'status', 'Admitted', 'fall_2017')

# total admitted count in Fall 2018:
admt_total_f18 = find_value(df_admissions_total, 'gender', 'Total', 'status', 'Admitted', 'fall_2018')

# total admitted count in Fall 2019:
admt_total_f19 = find_value(df_admissions_total, 'gender', 'Total', 'status', 'Admitted', 'fall_2019')

# total admitted count in Fall 2020:
admt_total_f20 = find_value(df_admissions_total, 'gender', 'Total', 'status', 'Admitted', 'fall_2020')

# total admitted count in Fall 2021:
admt_total_f21 = find_value(df_admissions_total, 'gender', 'Total', 'status', 'Admitted', 'fall_2021')



# ~~~~~ CALCULATED ADMISSION RATES [ADMITTED/APPLIED] ~~~~~~

# CALCULATED ADMISSION RATE FOR MEN BY YEAR

# men calculated admission rate in Fall 2017:
a_rate_men_f17 = admt_men_f17/app_men_f17

# men calculated admission rate in Fall 2018:
a_rate_men_f18 = admt_men_f18/app_men_f18

# men calculated admission rate in Fall 2019:
a_rate_men_f19 = admt_men_f19/app_men_f19

# men calculated admission rate in Fall 2020:
a_rate_men_f20 = admt_men_f20/app_men_f20

# men calculated admission rate in Fall 2021:
a_rate_men_f21 = admt_men_f21/app_men_f21



# CALCULATED ADMISSION RATE FOR WOMEN BY YEAR

# women calculated admission rate in Fall 2017:
a_rate_women_f17 = admt_women_f17/app_women_f17

# women calculated admission rate in Fall 2018:
a_rate_women_f18 = admt_women_f18/app_women_f18

# women calculated admission rate in Fall 2019:
a_rate_women_f19 = admt_women_f19/app_women_f19

# women calculated admission rate in Fall 2020:
a_rate_women_f20 = admt_women_f20/app_women_f20

# women calculated admission rate in Fall 2021:
a_rate_women_f21 = admt_women_f21/app_women_f21



# CALCULATED ADMISSION RATE FOR "NOT REPORTED" BY YEAR

# "not reported" calculated admission rate in Fall 2017:
a_rate_notreported_f17 = admt_notreported_f17/app_notreported_f17

# "not reported" calculated admission rate in Fall 2018:
a_rate_notreported_f18 = admt_notreported_f18/app_notreported_f18

# "not reported" calculated admission rate in Fall 2019:
a_rate_notreported_f19 = admt_notreported_f19/app_notreported_f19

# "not reported" calculated admission rate in Fall 2020:
a_rate_notreported_f20 = admt_notreported_f20/app_notreported_f20

# "not reported" calculated admission rate in Fall 2021:
a_rate_notreported_f21 = admt_notreported_f21/app_notreported_f21



# CALCULATED ADMISSION RATE TOTAL BY YEAR

# total calculated admission rate in Fall 2017:
a_rate_total_f17 = admt_total_f17/app_total_f17

# total calculated admission rate in Fall 2018:
a_rate_total_f18 = admt_total_f18/app_total_f18

# total calculated admission rate in Fall 2019:
a_rate_total_f19 = admt_total_f19/app_total_f19

# total calculated admission rate in Fall 2020:
a_rate_total_f20 = admt_total_f20/app_total_f20

# total calculated admission rate in Fall 2021:
a_rate_total_f21 = admt_total_f21/app_total_f21



# ~~~~~~~~~~~ ADMISSION RATE DATAFRAME ~~~~~~~~~~~~~


df_admission_rate = pd.DataFrame([
    ['Admission Rate', 'Men', a_rate_men_f17, a_rate_men_f18, a_rate_men_f19, a_rate_men_f20, a_rate_men_f21],
    ['Admission Rate', 'Women', a_rate_women_f17, a_rate_women_f18, a_rate_women_f19, a_rate_women_f20, a_rate_women_f21],
    ['Admission Rate', 'Not Reported', a_rate_notreported_f17, a_rate_notreported_f18,
                        a_rate_notreported_f19, a_rate_notreported_f20, a_rate_notreported_f21],
    ['Admission Rate', 'Total', a_rate_total_f17, a_rate_total_f18, a_rate_total_f19, a_rate_total_f20, a_rate_total_f21]],
                   columns=rates_headers)



# ========================================== ENROLLMENT RATE ===========================================================



# CREATE CALCULATED VALUES FOR "ENROLLMENT_RATE" DF (WHICH WILL BE CONCATENATED TO ORIGINAL ADMISSIONS DF LATER)


# ~~~~~~~~~~~~~~~ ENROLLMENT COUNTS ~~~~~~~~~~~~~~~


# ENROLLMENT COUNT FOR MEN BY YEAR

# men enrollment count in Fall 2017:
enr_men_f17 = find_value(df_admissions_total, 'gender', 'Men', 'status', 'Enrolled', 'fall_2017')

# men enrollment count in Fall 2018:
enr_men_f18 = find_value(df_admissions_total, 'gender', 'Men', 'status', 'Enrolled', 'fall_2018')

# men enrollment count in Fall 2019:
enr_men_f19 = find_value(df_admissions_total, 'gender', 'Men', 'status', 'Enrolled', 'fall_2019')

# men enrollment count in Fall 2020:
enr_men_f20 = find_value(df_admissions_total, 'gender', 'Men', 'status', 'Enrolled', 'fall_2020')

# men enrollment count in Fall 2021:
enr_men_f21 = find_value(df_admissions_total, 'gender', 'Men', 'status', 'Enrolled', 'fall_2021')



# ENROLLMENT COUNT FOR WOMEN BY YEAR

# women enrollment count in Fall 2017:
enr_women_f17 = find_value(df_admissions_total, 'gender', 'Women', 'status', 'Enrolled', 'fall_2017')

# women enrollment count in Fall 2018:
enr_women_f18 = find_value(df_admissions_total, 'gender', 'Women', 'status', 'Enrolled', 'fall_2018')

# women enrollment count in Fall 2019:
enr_women_f19 = find_value(df_admissions_total, 'gender', 'Women', 'status', 'Enrolled', 'fall_2019')

# women enrollment count in Fall 2020:
enr_women_f20 = find_value(df_admissions_total, 'gender', 'Women', 'status', 'Enrolled', 'fall_2020')

# women enrollment count in Fall 2021:
enr_women_f21 = find_value(df_admissions_total, 'gender', 'Women', 'status', 'Enrolled', 'fall_2021')



# ENROLLMENT COUNT TOTAL BY YEAR

# total enrollment count in Fall 2017:
enr_total_f17 = find_value(df_admissions_total, 'gender', 'Total', 'status', 'Enrolled', 'fall_2017')

# total enrollment count in Fall 2018:
enr_total_f18 = find_value(df_admissions_total, 'gender', 'Total', 'status', 'Enrolled', 'fall_2018')

# total enrollment count in Fall 2019:
enr_total_f19 = find_value(df_admissions_total, 'gender', 'Total', 'status', 'Enrolled', 'fall_2019')

# total enrollment count in Fall 2020:
enr_total_f20 = find_value(df_admissions_total, 'gender', 'Total', 'status', 'Enrolled', 'fall_2020')

# total enrollment count in Fall 2021:
enr_total_f21 = find_value(df_admissions_total, 'gender', 'Total', 'status', 'Enrolled', 'fall_2021')



# ~~~~~ CALCULATED ENROLLMENT RATES [ENROLLED/ADMITTED] ~~~~~~


# CALCULATED ENROLLMENT RATE FOR MEN BY YEAR

# men calculated enrollment rate in Fall 2017:
e_rate_men_f17 = enr_men_f17/admt_men_f17

# men calculated enrollment rate in Fall 2018:
e_rate_men_f18 = enr_men_f18/admt_men_f18

# men calculated enrollment rate in Fall 2019:
e_rate_men_f19 = enr_men_f19/admt_men_f19

# men calculated enrollment rate in Fall 2020:
e_rate_men_f20 = enr_men_f20/admt_men_f20

# men calculated enrollment rate in Fall 2021:
e_rate_men_f21 = enr_men_f21/admt_men_f21



# CALCULATED ENROLLMENT RATE FOR WOMEN BY YEAR

# women calculated enrollment rate in Fall 2017:
e_rate_women_f17 = enr_women_f17/admt_women_f17

# women calculated enrollment rate in Fall 2018:
e_rate_women_f18 = enr_women_f18/admt_women_f18

# women calculated enrollment rate in Fall 2019:
e_rate_women_f19 = enr_women_f19/admt_women_f19

# women calculated enrollment rate in Fall 2020:
e_rate_women_f20 = enr_women_f20/admt_women_f20

# women calculated enrollment rate in Fall 2021:
e_rate_women_f21 = enr_women_f21/admt_women_f21



# CALCULATED ENROLLMENT RATE TOTAL BY YEAR

# total calculated enrollment rate in Fall 2017:
e_rate_total_f17 = enr_total_f17/admt_total_f17

# total calculated enrollment rate in Fall 2018:
e_rate_total_f18 = enr_total_f18/admt_total_f18

# total calculated enrollment rate in Fall 2019:
e_rate_total_f19 = enr_total_f19/admt_total_f19

# total calculated enrollment rate in Fall 2020:
e_rate_total_f20 = enr_total_f20/admt_total_f20

# total calculated enrollment rate in Fall 2021:
e_rate_total_f21 = enr_total_f21/admt_total_f21



# ~~~~~~~~~~~ ENROLLMENT RATE DATAFRAME ~~~~~~~~~~~~~


df_enrollment_rate = pd.DataFrame([
    ['Enrollment Rate', 'Men', e_rate_men_f17, e_rate_men_f18, e_rate_men_f19, e_rate_men_f20, e_rate_men_f21],
    ['Enrollment Rate', 'Women', e_rate_women_f17, e_rate_women_f18, e_rate_women_f19, e_rate_women_f20, e_rate_women_f21],
    ['Enrollment Rate', 'Total', e_rate_total_f17, e_rate_total_f18, e_rate_total_f19, e_rate_total_f20, e_rate_total_f21]],
                   columns=rates_headers)



# ============================================ RATES TOTAL ===========================================================



# CONCATENATE "ADMISSION_RATE" DATAFRAME & "ENROLLMENT RATE" DATAFRAME TO CREATE "RATES_TOTAL" DATAFRAME

# create "frames" list of dataframes to use in concat formula:
frames = [df_admission_rate, df_enrollment_rate]

# concat dataframes using "frames" variable and resulting in rates_total dataframe
df_rates_total = pd.concat(frames)

# print check:
# print(df_rates_total)



# ======================================== WRITE TO EXCEL ===========================================================



# WRITE ALL DATAFRAMES TO EXCEL FILE
with pd.ExcelWriter(write_data_file_name) as writer:
     df_retention_gender_only.to_excel(writer, sheet_name='retention_gender_only', index=False)
     df_retention_gender_race.to_excel(writer, sheet_name='retention_gender_race', index=False)
     df_retention_race_only.to_excel(writer, sheet_name='retention_race_only', index=False)
     df_retention_total.to_excel(writer, sheet_name='retention_total', index=False)
     df_grad_gender_only.to_excel(writer, sheet_name='grad_gender_only', index=False)
     df_grad_gender_race.to_excel(writer, sheet_name='grad_gender_race', index=False)
     df_grad_race_only.to_excel(writer, sheet_name='grad_race_only', index=False)
     df_grad_total.to_excel(writer, sheet_name='grad_total', index=False)
     df_admissions_total.to_excel(writer, sheet_name='admissions_total', index=False)
     df_rates_total.to_excel(writer, sheet_name='rates_total', index=False)



# ======================================== END OF FILE ===========================================================
