# =====================================================================
# MSNI 2026 LEBANON - FULL SCRIPT
# Rules  : REACH_MSNA_2026_MSNI_final.xlsx (all tabs)
# Package: humind 2026.4.1
# Output : outputs/MSNI_2026_LBN_results_v6.xlsx (styled) + outputs/msni_2026_lbn_v6.rds
#          outputs/MSNI_Analysis_Table_MSNA2026_v6.xlsx (REACH format)
# =====================================================================
# Sections
#  0. Setup                 6. Protection
#  1. Rules (workbook)      7. WASH
#  2. Load data             8. Health
#  3. Food Security         9. MSNI + number of sectors in need + needs profiles
#  4. Education            10. Disaggregation variables
#  5. SNFI                 11. Result tables (+ 11b. People in need estimates)
#                          12. Export (styled Excel: colours, comments, algorithm)
#                          13. Analysis table (REACH format, one file, pop_group column)
# =====================================================================

# =====================================================================
# 0. SETUP
# =====================================================================
# install.packages(c("openxlsx", "analysistools", "srvyr"))   # run once
library(humind)
library(dplyr)
library(readxl)
library(openxlsx)
library(analysistools)
library(srvyr)

data_file   <- "REACH_LBN_Clean_Dataset_MSNA_2026_recoded_v2.xlsx"
pop_input   <- "D:/Lebanon/MSNA_2026/Population/05. 2027 LRP - Population Package.xlsx"   # 2027 LRP population
pop_sheet   <- "Population2027"
pal_sheet   <- "PALESTINIAN"                 # PRL / PRS camp population by district
id_col_main <- "uuid"
id_col_loop <- "_submission__uuid"
dir.create("outputs", showWarnings = FALSE)
out_suffix  <- "_v6"                      # added to every output file name (v1-v5 = shared versions)

# =====================================================================
# 1. RULES FROM THE MSNI WORKBOOK (edit here only)
# =====================================================================

# ---- Food Security tab ----
# LCS-FS: 10 standard strategies (humind). LCS-EN not used.
# Workbook, Food Security Step 2: LCSI reflects only strategies USED in the past 30 days
lcsi_var <- "fsl_lcsi_cat_yes"

# ---- Education tab ----
edu_age_start <- 5
edu_age_end   <- 17                                   # humind default 5-17 (comparability)
edu_protection_barriers <- c(                         # severity 4
  "protection_at_school", "protection_travel_school", "child_work_home",
  "child_work_outside", "child_armed_group", "child_marriage", "child_pregnancy",
  "restriction_or_ban", "enroll_lack_documentation", "discrimination")
edu_other_barriers <- c(                              # severity 3 (+ flags below)
  "costs", "lack_interest", "not_priority", "lack_accessible_school",
  "lack_classrooms", "lack_wash_facilities", "school_closed", "lack_teacher",
  "curriculum_not_useful", "child_health", "language", "enroll_displacement",
  "attending_a_non_formal_education",
  "child_too_young", "child_graduated", "not_school_age")
edu_undefined_barriers <- c("other", "dnk", "pnta", "ban")   # "ban" removed from tool
edu_flag_barriers      <- c("child_too_young", "child_graduated", "not_school_age")  # Flag / 1
edu_attack_score       <- 3                           # disruption by attack = 3 (humind = 4)

# ---- SNFI tab ----
snfi_inadequate_situation <- c("collective_center", "hosting", "hosted")
snfi_adequate   <- c("house", "apartment", "prefab_unit", "seperate_room_shared_apt", "studio")
snfi_inadequate <- c("unfinished_building", "tent", "makeshift", "apartment_house_extension",
                     "active_construction_site", "agricultural_engine_pump_room",
                     "construction_site", "concierge_s_room_in_residential_building",
                     "hotel_room", "room_at_workplace", "shop")
snfi_undefined  <- c("other", "dnk", "pnta")
occ_low    <- c("ownership")
occ_medium <- c("rented", "hosted_free", "sub_rented", "shared_rental_with_another_household",
                "hosted_by_relatives_or_friends", "provided_by_an_employer",
                "provided_by_organization", "caretaker_or_concierge_arrangement")
occ_high   <- c("no_agreement", "occupied_without_permission")
occ_undef  <- c("other", "dnk", "pnta")
damage_extra_undefined <- c(                          # Lebanon options = Undefined
  "foundation_or_column_damage_cracks", "partial_collapse_or_risk_of_collapse",
  "exposed_electrical_wiring", "visible_leakage_or_dampness", "non_functional_latrines",
  "non_functional_sewage_network", "non_functional_water_network")

# ---- Protection tab ----
prot_farm_weight <- 2                                 # yes_farm = 2 (humind = 1)

# ---- WASH tab ----
water_improved <- c("piped_dwelling", "piped_compound", "piped_neighbour", "tap", "borehole",
                    "protected_well", "well_spring", "sabeel_water", "sachet_water",
                    "rainwater_collection", "tank_truck", "tank_truck_free", "cart_tank",
                    "kiosk", "bottled_water")
water_undefined <- c("other", "dnk", "pnta")


# Sanitation (WASH tab, read by label; twin pits not listed -> humind default)
san_improved   <- c("flush_piped_sewer", "flush_septic_tank", "flush_pit_latrine",
                    "flush_dnk_where", "pit_latrine_slab", "compost",
                    "twin_pit_latrine_slab")
san_unimproved <- c("flush_open_drain", "pit_latrine_wo_slab", "bucket", "hanging_toilet",
                    "other_container", "twin_pit_latrine_wo_slab")
san_none       <- "none"
san_undefined  <- c("other", "dnk", "pnta")


# ---- Health tab: no deviation ----

# ---- Disaggregation tab ----
hoh_age_breaks <- c(-Inf, 18, 60, Inf)
hoh_age_labels <- c("Below 18", "18-59", "60 and above")
income_breaks  <- c(-Inf, 0.2, 0.4, 0.6, 0.8, Inf)
income_labels  <- c("Below 20% of median income", "20%-40% of median income",
                    "40%-60% of median income", "60%-80% of median income",
                    "Above 80% of median income")
min_n <- 30                                           # flag estimates below 30 HH


# =====================================================================
# 2. LOAD DATA
# =====================================================================
main_raw <- read_excel(data_file, sheet = "hh_clean_data",     guess_max = 10000)
edu_raw  <- read_excel(data_file, sheet = "edu_clean_data",    guess_max = 10000)
hlt_raw  <- read_excel(data_file, sheet = "health_clean_data", guess_max = 10000)

# remove indicators already calculated in the clean dataset (humind recalculates them)
precalc <- c(
  "fsl_fcs_score", "fsl_fcs_cat", "fsl_fcs_cat_ipc",
  grep("^fsl_hhs_.*(_recoded|_comp\\d|_score|_cat|_cat_ipc)$", names(main_raw), value = TRUE),
  grep("^fsl_lcsi_(stress|crisis|emergency)(_yes|_exhaust)?$", names(main_raw), value = TRUE),
  "fsl_lcsi_cat_yes", "fsl_lcsi_cat_exhaust", "fsl_lcsi_cat",
  "fsl_rcsi_score", "fsl_rcsi_cat")

main <- main_raw |>
  select(-any_of(precalc), -starts_with("hwise_score")) |>
  mutate(start = as.Date(start), weight = as.numeric(weight))

cat("Households:", nrow(main), "| Children:", nrow(edu_raw), "| Individuals:", nrow(hlt_raw), "\n")

# =====================================================================
# 3. FOOD SECURITY
# =====================================================================
main <- main |>
  add_lcsi() |>
  add_fcs(cutoffs = "normal") |>
  add_hhs() |>
  add_rcsi() |>
  add_fcm_phase() |>
  add_fclcm_phase(lcs_cat_var = lcsi_var) |>
  add_comp_foodsec()

table(main$comp_foodsec_score, useNA = "ifany")

# =====================================================================
# 4. EDUCATION
# =====================================================================
edu_ind <- edu_raw |>
  add_loop_edu_ind_age_corrected(
    main = main, id_col_loop = id_col_loop, id_col_main = id_col_main,
    ind_age = "edu_ind_age",
    schooling_start_age = edu_age_start, schooling_end_age = edu_age_end) |>
  add_loop_edu_access_d() |>
  add_loop_edu_barrier_protection_d(
    protection_issues     = edu_protection_barriers,
    non_protection_issues = edu_other_barriers,
    barriers_undefined    = edu_undefined_barriers) |>
  add_loop_edu_disrupted_d()

# Flag / 1: child too young / graduated / not school age -> not counted as "no access"
edu_flagged <- edu_ind$edu_ind_age_schooling %in% 1 &
  edu_ind$edu_barrier %in% edu_flag_barriers &
  edu_ind$edu_ind_no_access_d %in% 1
edu_ind$edu_ind_no_access_d[edu_flagged] <- 0
cat("School-age children:", sum(edu_ind$edu_ind_age_schooling %in% 1),
    "| Flag/1 children:", sum(edu_flagged), "\n")

main <- main |>
  add_loop_edu_ind_schooling_age_d_to_main(loop = edu_ind, id_col_main = id_col_main, id_col_loop = id_col_loop) |>
  add_loop_edu_access_d_to_main(loop = edu_ind, id_col_main = id_col_main, id_col_loop = id_col_loop) |>
  add_loop_edu_barrier_protection_d_to_main(loop = edu_ind, id_col_main = id_col_main, id_col_loop = id_col_loop) |>
  add_loop_edu_disrupted_d_to_main(loop = edu_ind, id_col_main = id_col_main, id_col_loop = id_col_loop) |>
  add_comp_edu() |>
  # disruption by attack = 3 (attack is the only path to 4 in the disruption dimension)
  mutate(
    comp_edu_score_disrupted = if_else(comp_edu_score_disrupted %in% 4,
                                       edu_attack_score, comp_edu_score_disrupted),
    comp_edu_score          = pmax(comp_edu_score_disrupted, comp_edu_score_attendance, na.rm = TRUE),
    comp_edu_in_need        = as.numeric(comp_edu_score >= 3),
    comp_edu_in_severe_need = as.numeric(comp_edu_score >= 4))

table(main$comp_edu_score, useNA = "ifany")

# =====================================================================
# 5. SNFI
# =====================================================================
# Lebanon damage options = Undefined: households reporting ONLY these options -> 'other'
dmg_extra_cols <- intersect(paste0("snfi_shelter_damage.", damage_extra_undefined), names(main))
dmg_std_cols   <- intersect(paste0("snfi_shelter_damage.",
                                   c("none", "minor_roof", "major_roof", "windows_doors", "floors",
                                     "walls", "total_collapse", "other", "dnk", "pnta")), names(main))
dmg_only_extra <- rowSums(main[, dmg_extra_cols], na.rm = TRUE) > 0 &
  rowSums(main[, dmg_std_cols],   na.rm = TRUE) == 0
main$snfi_shelter_damage.other[dmg_only_extra] <- 1
cat("HH with only Lebanon damage options (Undefined):", sum(dmg_only_extra), "\n")

main <- main |>
  add_shelter_type_cat(
    shelter_type = "snfi_shelter_type", sl_collective_center = snfi_inadequate_situation,
    adequate = snfi_adequate, inadequate = snfi_inadequate, undefined = snfi_undefined) |>
  add_shelter_issue_cat(sep = ".") |>
  add_shelter_damage_cat(sep = ".") |>
  add_fds_cannot_cat() |>
  add_occupancy_cat(
    occupancy_high_risk = occ_high, occupancy_medium_risk = occ_medium,
    occupancy_low_risk = occ_low, occupancy_undefined = occ_undef) |>
  add_comp_snfi()

table(main$comp_snfi_score, useNA = "ifany")

# =====================================================================
# 6. PROTECTION
# =====================================================================
main <- main |>
  add_prot_score_movement(sep = ".", other_safety_measures = "other") |>
  add_prot_score_practices(sep = ".") |>
  add_prot_score_rights(sep = ".") |>
  add_comp_prot()

# yes_farm weight 2: humind uses 1 -> add the difference and recalculate practices
main <- main |>
  mutate(
    .farm = if_else(prot_needs_2_activities.yes_farm %in% 1, prot_farm_weight - 1, 0),
    comp_prot_score_prot_needs_2_activities = comp_prot_score_prot_needs_2_activities + .farm,
    .tot = if_else(is.na(comp_prot_score_prot_needs_2_activities) &
                     is.na(comp_prot_score_prot_needs_2_social), NA_real_,
                   rowSums(cbind(comp_prot_score_prot_needs_2_activities,
                                 comp_prot_score_prot_needs_2_social), na.rm = TRUE)),
    comp_prot_score_practices = case_when(
      is.na(.tot) ~ NA_real_, .tot == 0 ~ 1, .tot == 1 ~ 2, .tot <= 3 ~ 3, TRUE ~ 4)) |>
  select(-.farm, -.tot, -comp_prot_score, -comp_prot_in_need, -comp_prot_in_severe_need) |>
  add_comp_prot()

table(main$comp_prot_score, useNA = "ifany")

# =====================================================================
# 7. WASH
# =====================================================================
main <- main |>
  add_hwise() |>
  add_drinking_water_source_cat(
    drinking_water_source_cat_improved  = water_improved,
    drinking_water_source_cat_undefined = water_undefined) |>
  add_drinking_water_time_cat() |>
  add_drinking_water_time_threshold_cat() |>
  add_drinking_water_quality_jmp_cat() |>
  add_sanitation_facility_cat(
    improved = san_improved, unimproved = san_unimproved,
    none = san_none, undefined = san_undefined) |>
  add_sharing_sanitation_facility_cat() |>
  add_sharing_sanitation_facility_n_ind() |>
  add_sanitation_facility_jmp_cat() |>
  add_handwashing_facility_cat() |>
  add_comp_wash()

table(main$comp_wash_score, useNA = "ifany")

# =====================================================================
# 8. HEALTH
# =====================================================================
hlt_ind <- hlt_raw |> add_loop_healthcare_needed_cat()

main <- main |>
  add_loop_healthcare_needed_cat_to_main(
    loop = hlt_ind, id_col_main = id_col_main, id_col_loop = id_col_loop) |>
  add_comp_health()

table(main$comp_health_score, useNA = "ifany")

# =====================================================================
# 9. MSNI + NUMBER OF SECTORS IN NEED
# =====================================================================
msni <- add_msni(main)

need_flags   <- c("comp_foodsec_in_need", "comp_edu_in_need", "comp_snfi_in_need",
                  "comp_prot_in_need", "comp_wash_in_need", "comp_health_in_need")
severe_flags <- c("comp_foodsec_in_severe_need", "comp_edu_in_severe_need",
                  "comp_snfi_in_severe_need", "comp_prot_in_severe_need",
                  "comp_wash_in_severe_need")            # Health max = 3

# recomputed so that households with no sector in need = 0 (humind leaves NA)
msni <- msni |>
  mutate(
    n_sectors_in_need        = if_else(is.na(msni_score), NA_real_,
                                       rowSums(across(all_of(need_flags)), na.rm = TRUE)),
    n_sectors_in_severe_need = if_else(is.na(msni_score), NA_real_,
                                       rowSums(across(all_of(severe_flags)), na.rm = TRUE)))

# ---- needs profiles (MSNI metric 4): which sectors are in need together ----
# order of sectors in the label: Food Security - Education - SNFI - Protection - WASH - Health
need_labels   <- c("Food Security", "Education", "SNFI", "Protection", "WASH", "Health")
severe_labels <- c("Food Security", "Education", "SNFI", "Protection", "WASH")     # Health max = 3
make_profile <- function(d, flags, labs, none) {
  m <- sapply(flags, function(f) d[[f]] %in% 1)
  if (is.null(dim(m))) m <- matrix(m, nrow = 1)
  apply(m, 1, function(r) if (any(r)) paste(labs[r], collapse = " - ") else none)
}
msni <- msni |>
  mutate(
    msni_needs_profile = if_else(is.na(msni_score), NA_character_,
                                 make_profile(msni, need_flags, need_labels, "No sectoral need")),
    msni_acute_profile = if_else(is.na(msni_score), NA_character_,
                                 make_profile(msni, severe_flags, severe_labels, "No acute sectoral need")))

table(msni$msni_score, useNA = "ifany")
table(msni$n_sectors_in_need, useNA = "ifany")
cat("Distinct needs profiles:", n_distinct(msni$msni_needs_profile, na.rm = TRUE),
    "| acute profiles:", n_distinct(msni$msni_acute_profile, na.rm = TRUE), "\n")
head(sort(table(msni$msni_needs_profile), decreasing = TRUE), 10)

# =====================================================================
# 10. DISAGGREGATION VARIABLES (Disaggregation tab + displacement + disability)
# =====================================================================

# ---- weighted median income (reference for income brackets) ----
wmedian <- function(x, w) {
  ok <- !is.na(x) & !is.na(w); x <- x[ok]; w <- w[ok]
  o <- order(x); x <- x[o]; w <- w[o]
  x[which(cumsum(w) >= sum(w) / 2)[1]]
}
income_median <- wmedian(msni$cm_income_total_lbp, msni$weight)          # across all pop groups
income_median_pop <- sapply(split(msni, msni$pop_group),                   # per pop group
                            function(d) wmedian(d$cm_income_total_lbp, d$weight))
if (any(income_median_pop <= 0, na.rm = TRUE))
  cat("WARNING: weighted median income is 0 for:", names(income_median_pop)[income_median_pop <= 0], "\n")

# ---- disaggregation groups (factors = logical order in the Excel) ----
msni <- msni |>
  mutate(
    # geography
    governorate    = admin1_label_en,
    district       = admin2_label_en,
    camp_name      = camp_label_en,
    
    # head of household gender (pnta excluded)
    hoh_gender_grp = factor(if_else(hoh_gender_final %in% c("female", "male"), hoh_gender_final, NA_character_),
                            levels = c("female", "male")),
    
    # head of household age: Below 18 / 18-59 / 60 and above
    hoh_age_grp    = cut(hoh_age_final, hoh_age_breaks, hoh_age_labels, right = FALSE),
    
    # income bracket: % of weighted median income (workbook categories)
    #   income_grp     = vs median of the household's OWN pop group (used within pop groups)
    #   income_grp_all = vs median across ALL pop groups      (used for "Overall" rows)
    income_grp     = cut(cm_income_total_lbp / income_median_pop[as.character(pop_group)],
                         income_breaks, income_labels, right = FALSE),
    income_grp_all = cut(cm_income_total_lbp / income_median, income_breaks, income_labels, right = FALSE),
    
    # displacement: Non-displaced / Returnee / IDP ("Refugee" 9 HH excluded)
    displacement_grp = factor(case_when(
      displacement_status %in% c("Non Displaced", "Non-displaced (economic/voluntary mover)") ~ "Non-displaced",
      displacement_status %in% c("Returnee IDP")                                             ~ "Returnee",
      displacement_status %in% c("IDP", "Infra-settlement IDP")                              ~ "IDP",
      TRUE ~ NA_character_), levels = c("Non-displaced", "Returnee", "IDP")),
    
    # disability: at least one household member with a disability
    disability_grp = factor(case_when(
      hh_with_disabled == "no"  ~ "No",
      hh_with_disabled == "yes" ~ "Yes - at least one member with disability",
      TRUE ~ NA_character_), levels = c("No", "Yes - at least one member with disability")))

# ---- checks ----
cat("Weighted median income, all pop groups (LBP):", format(income_median, big.mark = ","), "\n")
print(format(income_median_pop, big.mark = ","))
table(msni$hoh_gender_grp,   useNA = "ifany")
table(msni$hoh_age_grp,      useNA = "ifany")
table(msni$pop_group, msni$income_grp, useNA = "ifany")
table(msni$income_grp_all,   useNA = "ifany")
table(msni$displacement_grp, useNA = "ifany")
table(msni$disability_grp,   useNA = "ifany")

# =====================================================================
# 11. RESULT TABLES (weighted, % of households)
# =====================================================================
sectors <- c("Food Security" = "comp_foodsec", "Education" = "comp_edu", "SNFI" = "comp_snfi",
             "Protection" = "comp_prot", "WASH" = "comp_wash", "Health" = "comp_health",
             "MSNI" = "msni")
sector_max <- c("Food Security" = 5, "Education" = 4, "SNFI" = 5, "Protection" = 4,
                "WASH" = 5, "Health" = 3, "MSNI" = 5)
sev_levels <- c("1", "2", "3", "4", "4+")

wpct <- function(cond, w) round(sum(w[cond]) / sum(w) * 100, 1)

# ---- severity 1 - 4+ for every sector and the MSNI ----
severity_rows <- function(d) {
  bind_rows(lapply(names(sectors), function(lab) {
    s <- d[[paste0(sectors[[lab]], "_score")]]; w <- d$weight
    ok <- !is.na(s) & !is.na(w); s <- s[ok]; w <- w[ok]
    if (length(s) == 0) return(NULL)
    lev <- ifelse(s >= 5, "4+", as.character(s))
    pct <- sapply(sev_levels, function(l) wpct(lev == l, w))
    pct[c(1, 2, 3, 4, 5) > sector_max[[lab]]] <- NA          # level not reachable
    data.frame(Sector = lab, n_HH = length(s), as.list(pct),
               `% in need (3+)`     = wpct(s >= 3, w),
               `% severe need (4+)` = if (sector_max[[lab]] >= 4) wpct(s >= 4, w) else NA,
               check.names = FALSE)
  }))
}

# ---- MSNI metrics: in need, severe, average and distribution of sectors ----
msni_rows <- function(d) {
  d <- d[!is.na(d$msni_score) & !is.na(d$weight), ]
  if (nrow(d) == 0) return(NULL)
  w <- d$weight; nn <- d$n_sectors_in_need; ns <- d$n_sectors_in_severe_need
  out <- data.frame(
    n_HH = nrow(d),
    `% HH in need (3+)`            = wpct(d$msni_score >= 3, w),
    `% HH in severe need (4+)`     = wpct(d$msni_score >= 4, w),
    `Avg # sectors in need`        = round(sum(w * nn) / sum(w), 2),
    `Avg # sectors in severe need` = round(sum(w * ns) / sum(w), 2),
    check.names = FALSE)
  for (k in 0:6) out[[paste0("% HH in need in ", k, " sectors")]]        <- wpct(nn == k, w)
  for (k in 0:5) out[[paste0("% HH in severe need in ", k, " sectors")]] <- wpct(ns == k, w)
  # top 3 needs profiles (3+), "No sectoral need" excluded; % of all HH in the group
  prof <- d$msni_needs_profile
  keep <- !is.na(prof) & prof != "No sectoral need"
  pw   <- if (any(keep)) sort(tapply(w[keep], prof[keep], sum), decreasing = TRUE) / sum(w) * 100 else numeric(0)
  for (i in 1:3) {
    out[[paste0("Top ", i, " need profile")]] <- if (length(pw) >= i) names(pw)[i] else NA_character_
    out[[paste0("Top ", i, " %")]]            <- if (length(pw) >= i) round(as.numeric(pw[i]), 1) else NA_real_
  }
  out
}

# ---- apply a function overall or by group(s), optionally for some pop groups ----
by_group <- function(df, fun, group_vars = NULL, pops = NULL) {
  d <- if (is.null(pops)) df else filter(df, pop_group %in% pops)
  if (is.null(group_vars)) return(cbind(Group = "Overall", fun(d)))
  d |>
    filter(if_all(all_of(group_vars), ~ !is.na(.x))) |>
    group_by(across(all_of(group_vars))) |>
    group_modify(~ fun(.x)) |>
    ungroup() |>
    mutate(Flag = if_else(n_HH < min_n, paste0("Low sample (<", min_n, " HH)"), ""))
}

# ---- disaggregation plan ----
# pop_group (primary); governorate (all pop groups); district (Lebanese);
# camp (PRS, PRL); HoH gender, HoH age, income, displacement, disability (by pop_group)
# ov = variable used for the "Overall" rows (all pop groups together); none for district / camp
# income: Overall rows use income_grp_all (median across all groups), pop rows use own-group median
plan <- list(
  Overall        = list(g = NULL,                                p = NULL),
  Pop_group      = list(g = "pop_group",                         p = NULL),
  Governorate    = list(g = c("pop_group", "governorate"),       p = NULL, ov = "governorate"),
  District       = list(g = c("pop_group", "district"),          p = c("lebanese")),
  Camp           = list(g = c("pop_group", "camp_name"),         p = c("prs", "prl")),
  HoH_gender     = list(g = c("pop_group", "hoh_gender_grp"),    p = NULL, ov = "hoh_gender_grp"),
  HoH_age        = list(g = c("pop_group", "hoh_age_grp"),       p = NULL, ov = "hoh_age_grp"),
  Income         = list(g = c("pop_group", "income_grp"),        p = NULL, ov = "income_grp_all"),
  Displacement   = list(g = c("pop_group", "displacement_grp"),  p = NULL, ov = "displacement_grp"),
  Disability     = list(g = c("pop_group", "disability_grp"),    p = NULL, ov = "disability_grp"))

# ---- needs profiles: top 10 per group + "Other combinations" (weighted %) ----
top_n_profiles <- 10
profile_rows <- function(var) function(d) {
  d <- d[!is.na(d[[var]]) & !is.na(d$weight), ]
  if (nrow(d) == 0) return(NULL)
  w   <- tapply(d$weight, d[[var]], sum)
  cnt <- table(d[[var]])[names(w)]
  o   <- order(-w); w <- w[o]; cnt <- cnt[o]
  k   <- min(top_n_profiles, length(w))
  prof <- names(w)[1:k]; pw <- as.numeric(w[1:k]); pn <- as.integer(cnt[1:k])
  if (length(w) > k) {
    prof <- c(prof, "Other combinations"); pw <- c(pw, sum(w[-(1:k)])); pn <- c(pn, sum(cnt[-(1:k)]))
  }
  data.frame(n_HH = nrow(d), Rank = c(seq_len(k), rep(NA, length(prof) - k)),
             Profile = prof, `# sectors` = ifelse(prof %in% c("Other combinations"), NA,
                                                  ifelse(grepl("^No ", prof), 0, lengths(strsplit(prof, " - ")))),
             `n_HH with profile` = pn, `%` = round(pw / sum(as.numeric(w)) * 100, 1), check.names = FALSE)
}

# ---- by group + "Overall" rows on top (all pop groups together) ----
by_group_ov <- function(df, fun, x) {
  res <- by_group(df, fun, x$g, x$p)
  if (is.null(x[["ov"]])) return(res)               # [[ ]] = exact name (x$ov would partly match other fields)
  tgt  <- x$g[2]
  d_ov <- df |> mutate(pop_group = "Overall")
  d_ov[[tgt]] <- df[[x[["ov"]]]]
  bind_rows(by_group(d_ov, fun, x$g, NULL), res)
}

msni_tables     <- lapply(plan, function(x) by_group_ov(msni, msni_rows,     x))
severity_tables <- lapply(plan, function(x) by_group_ov(msni, severity_rows, x))
profile_tables  <- lapply(plan, function(x) by_group_ov(msni, profile_rows("msni_needs_profile"), x))
acute_tables    <- lapply(plan, function(x) by_group_ov(msni, profile_rows("msni_acute_profile"), x))

# quick view
print(msni_tables$Pop_group[, 1:6])
print(msni_tables$Displacement[, 1:5])
print(msni_tables$Disability[, 1:5])
print(profile_tables$Overall)

# ---- dimensions underlying each sector (weighted % per category, by pop_group) ----
dimension_vars <- c(
  "fsl_fcs_cat", "fsl_rcsi_cat", "fsl_hhs_cat_ipc", lcsi_var, "fsl_fc_phase", "fclcm_phase",
  "comp_edu_score_attendance", "comp_edu_score_disrupted",
  "comp_snfi_score_shelter_type_cat", "comp_snfi_score_shelter_issue_cat",
  "comp_snfi_score_tenure_security_cat", "comp_snfi_score_fds_cannot_cat",
  "comp_snfi_score_shelter_damage_cat",
  "comp_prot_score_movement", "comp_prot_score_practices", "comp_prot_score_rights",
  "comp_wash_score_water_quantity", "comp_wash_score_water_quality",
  "comp_wash_score_sanitation", "comp_wash_score_hygiene")
dimension_vars <- intersect(dimension_vars, names(msni))

dimensions <- bind_rows(lapply(dimension_vars, function(v) {
  bind_rows(
    msni |> mutate(pop_group = "Overall"),
    msni) |>
    filter(!is.na(.data[[v]]), !is.na(weight)) |>
    group_by(Group = pop_group, Category = as.character(.data[[v]])) |>
    summarise(w = sum(weight), n_HH = n(), .groups = "drop_last") |>
    mutate(Indicator = v, `%` = round(w / sum(w) * 100, 1)) |>
    ungroup() |>
    select(Indicator, Group, Category, n_HH, `%`)
}))

# ---- Protection sanity check (workbook): HH at severity 4 with only ONE option ----
n_sel <- function(q, opts) rowSums(msni[, intersect(paste0(q, ".", opts), names(msni))], na.rm = TRUE)

n_mov <- n_sel("prot_needs_3_movement", c("no_changes_feel_unsafe", "women_girls_avoid_places",
                                          "men_avoid_places", "boys_avoid_places", "women_girls_avoid_night", "men_avoid_night",
                                          "boys_avoid_night", "girls_boys_avoid_school", "different_routes", "avoid_markets",
                                          "avoid_public_offices", "avoid_fields", "women_girls_boys_avoid_firewood", "women_girls_boys_avoid_places"))
n_pra <- n_sel("prot_needs_2_activities", c("yes_work", "yes_livelihood", "yes_safety", "yes_farm",
                                            "yes_water", "yes_free_choices")) +
  n_sel("prot_needs_2_social", c("yes_visiting_family", "yes_visiting_friends", "yes_community_events",
                                 "yes_joining_groups", "yes_child_recreation", "yes_decision_making"))
n_rig <- n_sel("prot_needs_1_services", c("yes_healthcare", "yes_schools", "yes_therapeutic_services",
                                          "yes_edu_facilities", "yes_social_services", "yes_gov_services")) +
  n_sel("prot_needs_1_justice", c("yes_identity_documents", "yes_counselling_legal",
                                  "yes_property_docs", "yes_gov_services", "yes_birth_certificates"))

# ---- checks ----
checks <- data.frame(
  Check = c("Households", "Food Security: HH with no score", paste0("Education: school-age children (", edu_age_start, "-", edu_age_end, ")"),
            "Education: Flag/1 children", "SNFI: HH with only Lebanon damage options (Undefined)",
            "Protection: HH with no score",
            "Protection sanity check: movement = 4 with only 1 option",
            "Protection sanity check: practices = 4 with only 1 option",
            "Protection sanity check: rights = 4 with only 1 option",
            "Health: individuals needed care, 'received' missing",
            "Health: HH with no score", "Income: HH with no income reported",
            "Displacement: HH excluded (Refugee / missing)",
            "Disability: HH with at least one member with disability",
            "Weighted median income (LBP)", "LCSI variable used",
            paste0("Weighted median income - ", names(income_median_pop), " (LBP)"),
            "Needs profiles (3+): number of distinct combinations",
            "Acute needs profiles (4+): number of distinct combinations"),
  Value = c(nrow(msni), sum(is.na(msni$comp_foodsec_score)), sum(edu_ind$edu_ind_age_schooling %in% 1),
            sum(edu_flagged), sum(dmg_only_extra), sum(is.na(msni$comp_prot_score)),
            sum(msni$comp_prot_score_movement  %in% 4 & n_mov == 1),
            sum(msni$comp_prot_score_practices %in% 4 & n_pra == 1),
            sum(msni$comp_prot_score_rights    %in% 4 & n_rig == 1),
            sum(hlt_ind$health_ind_healthcare_needed %in% "yes" & is.na(hlt_ind$health_ind_healthcare_received)),
            sum(is.na(msni$comp_health_score)), sum(is.na(msni$cm_income_total_lbp)),
            sum(is.na(msni$displacement_grp)),
            sum(msni$disability_grp %in% "Yes - at least one member with disability"),
            format(income_median, big.mark = ","), lcsi_var,
            format(income_median_pop, big.mark = ","),
            n_distinct(msni$msni_needs_profile, na.rm = TRUE),
            n_distinct(msni$msni_acute_profile, na.rm = TRUE)))
print(checks)

# =====================================================================
# 11b. PEOPLE IN NEED - rough estimates (2027 LRP population, Syrians excluded)
#      # people = population x % of PEOPLE in need (households weighted by weight x hh_size)
#      Lebanese: by district (+ governorate = sum of districts); migrants: by governorate;
#      PRL / PRS: by governorate, CAMP population only (survey covered camps only;
#      PRL in gatherings / out of camps excluded)
# =====================================================================
pop27 <- read_excel(pop_input, sheet = pop_sheet) |>
  filter(!is.na(Governorate), !is.na(District), District != "Total") |>      # drop Total row
  transmute(governorate = if_else(Governorate == "Nabatieh", "El Nabatieh", Governorate),   # match dataset name
            district = District,
            lebanese = `TOTAL LEBANESE (2027)`,
            migrant  = `TOTAL MIGRANTS (2027)`,
            prl      = `Palestinian Refugees in Lebanon (PRL) (2027)`,
            prs      = `Palestinian Refugees from Syria (PRS) (2027)`)        # Syrians not used
# ---- PRL / PRS: replace total population by CAMP population (sheet PALESTINIAN) ----
pal_raw <- read_excel(pop_input, sheet = pal_sheet, col_names = FALSE, .name_repair = "minimal")
hdr     <- which(apply(pal_raw, 1, function(r) any(trimws(as.character(r)) == "District", na.rm = TRUE)))[1]
if (is.na(hdr)) stop("No 'District' header found in sheet ", pal_sheet)
pal <- pal_raw[-(1:hdr), ]                                   # table starts below the header row
nm  <- trimws(as.character(unlist(pal_raw[hdr, ])))
nm[is.na(nm) | nm == ""] <- paste0("empty_", which(is.na(nm) | nm == ""))
names(pal) <- make.unique(nm)
cat("PALESTINIAN header row:", hdr, "| columns:", paste(names(pal)[!grepl("^empty_", names(pal))], collapse = " | "), "\n")

col_prl <- grep("^PRL.*Camps", names(pal), value = TRUE, ignore.case = TRUE)[1]   # "PRL - Camps Population"
col_prs <- grep("^PRS.*Camps", names(pal), value = TRUE, ignore.case = TRUE)[1]   # "PRS in Camps"
if (is.na(col_prl) || is.na(col_prs)) stop("PRL / PRS camp columns not found in sheet ", pal_sheet)
cat("Using columns:", col_prl, "|", col_prs, "\n")

num0 <- function(x) { x <- suppressWarnings(as.numeric(x)); if_else(is.na(x), 0, x) }   # "-" -> 0
pal <- pal |>
  mutate(District = trimws(as.character(District))) |>
  filter(!is.na(District), District != "", !grepl("total", District, ignore.case = TRUE)) |>
  transmute(district = District,
            prl_camp = num0(.data[[col_prl]]),
            prs_camp = num0(.data[[col_prs]]))
pop27 <- pop27 |>
  left_join(pal, by = "district") |>
  mutate(prl = coalesce(prl_camp, 0), prs = coalesce(prs_camp, 0)) |>
  select(-prl_camp, -prs_camp)
cat("Districts in PALESTINIAN sheet not in population sheet:", setdiff(pal$district, pop27$district), "\n")

cat("\nPopulation 2027 (Syrians excluded; PRL / PRS = camps only):\n")
print(format(round(colSums(pop27[, c("lebanese", "migrant", "prl", "prs")])), big.mark = ","))

# ---- name check: population file vs dataset (should all be empty) ----
d_leb <- unique(msni$district[msni$pop_group == "lebanese"])
g_mig <- unique(msni$governorate[msni$pop_group == "migrant"])
cat("Districts in population file, not in Lebanese data :", setdiff(pop27$district, d_leb), "\n")
cat("Districts in Lebanese data, not in population file :", setdiff(d_leb, pop27$district), "\n")
cat("Governorates in population file, not in migrant data:", setdiff(unique(pop27$governorate), g_mig), "\n")
cat("Governorates in migrant data, not in population file:", setdiff(g_mig, unique(pop27$governorate)), "\n")

pin_level <- c(lebanese = "district", migrant = "governorate", prl = "governorate", prs = "governorate")   # PRL/PRS: camps by governorate
msni$hh_size <- as.numeric(msni$hh_size)

# % of people with score >= k (household weight x household size)
ppct <- function(d, s, k) {
  ok <- !is.na(d[[s]]) & !is.na(d$weight) & !is.na(d$hh_size) & d$hh_size > 0
  pw <- d$weight[ok] * d$hh_size[ok]
  if (length(pw) == 0 || sum(pw) == 0) return(NA_real_)
  sum(pw[d[[s]][ok] >= k]) / sum(pw) * 100
}

# ---- unit level: district (Lebanese), governorate (migrants), whole group (PRL, PRS) ----
pin_list <- list()
for (p in names(pin_level)) {
  dp <- filter(msni, pop_group == p)
  lv <- pin_level[[p]]
  pu <- if (is.na(lv)) tibble(unit = "All", population = sum(pop27[[p]])) else
    pop27 |> group_by(unit = .data[[lv]]) |> summarise(population = sum(.data[[p]]), .groups = "drop")
  pu <- filter(pu, population > 0)
  for (lab in names(sectors)) {
    s    <- paste0(sectors[[lab]], "_score")
    has4 <- sector_max[[lab]] >= 4
    nat3 <- ppct(dp, s, 3); nat4 <- if (has4) ppct(dp, s, 4) else NA_real_
    for (i in seq_len(nrow(pu))) {
      du   <- if (is.na(lv)) dp else dp[dp[[lv]] %in% pu$unit[i], ]
      n_hh <- sum(!is.na(du[[s]]))
      fb   <- n_hh == 0                                   # no survey data -> group %
      pin_list[[length(pin_list) + 1]] <- data.frame(
        Sector = lab, pop_group = p, Level = ifelse(is.na(lv), "population group", lv),
        Unit = pu$unit[i], n_HH = n_hh, `Population 2027` = pu$population[i],
        `% people 3+` = if (fb) nat3 else ppct(du, s, 3),
        `% people 4+` = if (!has4) NA_real_ else if (fb) nat4 else ppct(du, s, 4),
        Note = ifelse(fb, "No survey data: population-group % used", ""), check.names = FALSE)
    }
  }
}
sec_order <- c("MSNI", setdiff(names(sectors), "MSNI"))
pop_order <- c("Overall", "lebanese", "migrant", "prl", "prs")
pin_units <- bind_rows(pin_list) |>
  mutate(`# people 3+` = `Population 2027` * `% people 3+` / 100,
         `# people 4+` = `Population 2027` * `% people 4+` / 100)

# ---- totals per pop group and Overall ----
sum_na <- function(x) if (all(is.na(x))) NA_real_ else sum(x, na.rm = TRUE)
pin_pop <- pin_units |>
  group_by(Sector, pop_group) |>
  summarise(Level = first(Level), n_HH = sum(n_HH), `Population 2027` = sum(`Population 2027`),
            `# people 3+` = sum_na(`# people 3+`), `# people 4+` = sum_na(`# people 4+`), .groups = "drop")
pin_tot <- pin_pop |>
  group_by(Sector) |>
  summarise(pop_group = "Overall", Level = "all population groups", n_HH = sum(n_HH),
            `Population 2027` = sum(`Population 2027`),
            `# people 3+` = sum_na(`# people 3+`), `# people 4+` = sum_na(`# people 4+`), .groups = "drop")
people_in_need <- bind_rows(pin_tot, pin_pop) |>
  mutate(`% people 3+` = round(`# people 3+` / `Population 2027` * 100, 1),
         `% people 4+` = round(`# people 4+` / `Population 2027` * 100, 1),
         across(c(`Population 2027`, `# people 3+`, `# people 4+`), round)) |>
  arrange(factor(Sector, sec_order), factor(pop_group, pop_order)) |>
  select(Sector, pop_group, Level, n_HH, `Population 2027`,
         `% people 3+`, `# people 3+`, `% people 4+`, `# people 4+`)
# ---- Lebanese admin1: sum of districts per governorate (added after totals -> no double count) ----
dist_gov <- distinct(pop27, district, governorate)
leb_gov <- pin_units |>
  filter(pop_group == "lebanese") |>
  left_join(dist_gov, by = c("Unit" = "district")) |>
  group_by(Sector, pop_group, Unit = governorate) |>
  summarise(Level = "governorate (sum of districts)", n_HH = sum(n_HH),
            `Population 2027` = sum(`Population 2027`),
            `# people 3+` = sum_na(`# people 3+`), `# people 4+` = sum_na(`# people 4+`), .groups = "drop") |>
  mutate(`% people 3+` = `# people 3+` / `Population 2027` * 100,
         `% people 4+` = `# people 4+` / `Population 2027` * 100, Note = "")
# ---- Overall per governorate = Lebanese (sum of districts) + migrants + PRL + PRS ----
ov_gov <- bind_rows(leb_gov, filter(pin_units, pop_group %in% c("migrant", "prl", "prs"))) |>
  group_by(Sector, Unit) |>
  summarise(pop_group = "Overall", Level = "governorate", n_HH = sum(n_HH, na.rm = TRUE),
            `Population 2027` = sum(`Population 2027`),
            `# people 3+` = sum_na(`# people 3+`), `# people 4+` = sum_na(`# people 4+`), .groups = "drop") |>
  mutate(`% people 3+` = `# people 3+` / `Population 2027` * 100,
         `% people 4+` = `# people 4+` / `Population 2027` * 100,
         Note = "Sum of the 4 population groups (PRL/PRS = camp population)")

# ---- national rows (Overall + each pop group), from the totals table ----
nat_rows <- people_in_need |> mutate(Level = "national", Unit = "Lebanon", Note = "")

lvl_order <- c("national", "governorate", "governorate (sum of districts)", "district")
pin_units <- bind_rows(nat_rows, ov_gov, leb_gov, pin_units) |>
  mutate(across(c(`% people 3+`, `% people 4+`), ~ round(.x, 1)),
         across(c(`Population 2027`, `# people 3+`, `# people 4+`), round)) |>
  arrange(factor(Sector, sec_order), factor(pop_group, pop_order), factor(Level, lvl_order), Unit) |>
  select(Sector, pop_group, Level, Unit, n_HH, `Population 2027`,
         `% people 3+`, `# people 3+`, `% people 4+`, `# people 4+`, Note)
cat("Units with no survey data (group % used):", sum(grepl("No survey data", pin_units$Note) & pin_units$Sector == "MSNI"), "\n")
print(filter(people_in_need, Sector == "MSNI"))

# =====================================================================
# 12. EXPORT - styled Excel (colours, comments, algorithm)
# =====================================================================
saveRDS(msni, paste0("outputs/msni_2026_lbn", out_suffix, ".rds"))

# ---- README ----
readme <- data.frame(
  Item = c("Rules", "Package", "Unit", "Weights", "In need / severe need", "4+",
           "Sectors in need", "Needs profile", "People in need", "LCSI", "Disaggregation", "Flag", "Colours", "Run date"),
  Description = c(
    "REACH_MSNA_2026_MSNI_final.xlsx (Lebanon rules, all sector tabs)",
    paste("humind", packageVersion("humind")),
    "Household: Education and Health rosters aggregated to household before scoring",
    "All % weighted with survey weight 'weight'",
    "In need = score 3 or more; severe need = score 4 or more",
    "4+ = humind score 5. Empty = level not reachable (Health max 3; Education, Protection max 4)",
    "Count of sectors with score 3+ (0-6) and 4+ (0-5, Health excluded); 0 when no sector in need",
    paste0("Rough estimates: 2027 LRP population (Syrians excluded) x % of PEOPLE in need (households weighted by weight x hh_size). ",
           "Lebanese by district, migrants by governorate, PRL/PRS by governorate using CAMP population only (sheet PALESTINIAN; ",
           "gatherings / out-of-camp PRL not covered by the survey, excluded). ",
           "Sheets People_in_need (totals) and People_in_need_units (detail: national + governorate for Overall and all groups, ",
           "district for Lebanese)."),
    paste0("Combination of sectors in need (3+), e.g. 'SNFI - Protection'; Profile_ sheets = top ", top_n_profiles,
           " per group + 'Other combinations'. Acute_Profile_ sheets = same with sectors at 4+ (Health excluded)"),
    paste0(lcsi_var, ifelse(lcsi_var == "fsl_lcsi_cat_yes",
                            " (strategies used only - workbook Food Security Step 2)",
                            " (used + exhausted - humind default)")),
    paste0("pop_group; governorate (all pop groups); district (Lebanese); camp (PRS, PRL); HoH gender; HoH age; ",
           "income bracket (% of weighted median of the household's own pop group: ",
           paste(names(income_median_pop), format(income_median_pop, big.mark = ","), collapse = ", "),
           " LBP; Overall rows: median across all groups = ", format(income_median, big.mark = ","), " LBP); ",
           "displacement (Non-displaced / Returnee / IDP; 'Refugee' 9 HH excluded); disability (at least one member)"),
    paste0("Low sample = fewer than ", min_n, " households (orange rows)"),
    paste("Severity headers: 1 pink, 2 blue, 3 rose, 4 red, 4+ dark red. Red headers = % in need / severe",
          "(cell colour white 0% -> dark red 100%). Dark blue = averages. Brown = distribution by number of sectors.",
          "Purple = top 3 needs profiles (3+, 'No sectoral need' excluded).",
          "Group columns: Overall grey (all pop groups together), Lebanese blue, migrants green, PRL purple, PRS beige.",
          "Hover headers for comments."),
    as.character(Sys.Date())))

# ---- Algorithm: method step by step, with results of this run ----
ov  <- severity_tables$Overall
nd  <- function(s) ov$`% in need (3+)`[ov$Sector == s]
sv  <- function(s) ov$`% severe need (4+)`[ov$Sector == s]
nsc <- function(v) format(sum(!is.na(msni[[v]])), big.mark = ",")
mo  <- msni_tables$Overall

algorithm <- data.frame(
  Step  = c("DATA", "D1", "D2", "D3",
            "FOOD SECURITY", "F1", "F2",
            "EDUCATION", "E1", "E2", "E3", "E4",
            "SNFI", "S1", "S2",
            "PROTECTION", "P1", "P2",
            "WASH", "W1", "W2",
            "HEALTH", "H1", "H2", "H3",
            "MSNI", "M1", "M2", "M3",
            "RESULTS", "R1", "R2", "R3"),
  Level = c("", "Household", "Individual", "Individual",
            "", "Household", "Household",
            "", "Individual", "Individual", "Individual -> Household", "Household",
            "", "Household", "Household",
            "", "Household", "Household",
            "", "Household", "Household",
            "", "Individual", "Individual -> Household", "Household",
            "", "Household", "Household", "Household",
            "", "Household -> Population", "Household -> Population", "Household -> Population"),
  `What happens` = c(
    "", "hh_clean_data: one row per household", "edu_clean_data: one row per child / young person",
    "health_clean_data: one row per household member",
    "", "FCS, rCSI, HHS, LCSI (10 strategies)", "Step 1 matrix (FC phase) then Step 2 matrix (FC x LCSI) = score 1-5",
    "", paste0("Correct age to school year; keep ages ", edu_age_start, "-", edu_age_end), "Flag each child: no access, protection barrier, disruption",
    "Count flagged children per household (household ID)", "Score = max(access & barriers, disruption)",
    "", "Score shelter type, issues, tenure, FDS + lighting, damage", "Score = max of 5 dimensions",
    "", "Sum weighted options: movement, safe practices, rights & services", "Totals -> severity 1-4; score = max of 3",
    "", "Water quantity (H-WISE), quality, sanitation, hygiene by setting", "Score = max of 4 dimensions",
    "", "Classify each person: no need / met / unmet", "Count persons per household",
    "Any unmet = 3; else any met = 2; else 1",
    "", "MSNI = highest score across the 6 sectors", "Number of sectors in need (3+) and severe need (4+)",
    "Needs profile = which sectors are in need together (3+); acute profile (4+)",
    "", "Weighted % = sum of weights of HH meeting condition / sum of weights of HH with a score x 100",
    "Same by pop_group, governorate, district, camp, HoH gender, HoH age, income, displacement, disability",
    "People in need = 2027 population x % of people (weight x hh_size) in need"),
  `Rule (final workbook)` = c(
    "", "Food Security, SNFI, Protection, WASH, weight, disaggregation", "Asked 3-24; linked by _submission__uuid = uuid",
    "Linked by _submission__uuid = uuid",
    "", "Standard thresholds", paste("FEWS NET matrices; LCSI variable:", lcsi_var, "(Step 2: used only)"),
    "", "humind default ages 5-17", "Protection barrier = 4; other = 3; non-formal = 3; too young / graduated / not school age = Flag/1",
    "Sum of child flags", "No access = 3; + protection barrier = 4; teacher = 2; hazards / displaced / attack = 3",
    "", "Hosting/hosted inadequate; Lebanon shelter types and occupancy mapped; Lebanon-only damage = Undefined", "",
    "", "Workbook weights; yes_farm = 2; 'other' = NA", "Movement 0/1/2/3+; practices & rights 0/1/2-3/4+",
    "", "sabeel_water, tank_truck_free = improved; sanitation by Kobo codes (15 options)", "",
    "", "needed yes + received no = unmet", "Sum by household ID", "Max severity 3",
    "", "", "0 when no sector in need",
    "Order: FS - Edu - SNFI - Prot - WASH - Health; none = 'No sectoral need'",
    "", "Survey weight; HH with no score excluded", "Governorate: all pop groups; district: Lebanese; camp: PRS & PRL",
    "Lebanese by district, migrants and PRL/PRS (camps only) by governorate; Syrians excluded"),
  `Result (this run)` = c(
    "", paste(format(nrow(msni), big.mark = ","), "households"),
    paste(format(nrow(edu_raw), big.mark = ","), "individuals"),
    paste(format(nrow(hlt_raw), big.mark = ","), "individuals"),
    "", paste(nsc("comp_foodsec_score"), "HH scored"),
    paste0("In need ", nd("Food Security"), "%; severe ", sv("Food Security"), "%"),
    "", paste(format(sum(edu_ind$edu_ind_age_schooling %in% 1), big.mark = ","), "school-age children"),
    paste(sum(edu_flagged), "Flag/1 children"), "",
    paste0("In need ", nd("Education"), "%; severe ", sv("Education"), "%"),
    "", paste(sum(dmg_only_extra), "HH with only Lebanon damage options = Undefined"),
    paste0("In need ", nd("SNFI"), "%; severe ", sv("SNFI"), "%"),
    "", "", paste0("In need ", nd("Protection"), "%; severe ", sv("Protection"), "%"),
    "", "", paste0("In need ", nd("WASH"), "%; severe ", sv("WASH"), "%"),
    "", paste(sum(hlt_ind$health_ind_healthcare_needed %in% "yes" & is.na(hlt_ind$health_ind_healthcare_received)),
              "persons needed care with no 'received' answer"), "",
    paste0("In need ", nd("Health"), "%; ", sum(is.na(msni$comp_health_score)), " HH no score"),
    "", paste0("MSNI in need ", mo$`% HH in need (3+)`, "%; severe ", mo$`% HH in severe need (4+)`, "%"),
    paste0("Avg ", mo$`Avg # sectors in need`, " sectors in need; ", mo$`% HH in need in 0 sectors`, "% HH in no sector"),
    paste0(n_distinct(msni$msni_needs_profile, na.rm = TRUE), " profiles; most common: ",
           profile_tables$Overall$Profile[1], " (", profile_tables$Overall$`%`[1], "%)"),
    "", "All % in this file", paste0("Groups < ", min_n, " HH flagged"),
    with(filter(people_in_need, Sector == "MSNI", pop_group == "Overall"),
         paste0(format(`# people 3+`, big.mark = ","), " people in need (3+); ",
                format(`# people 4+`, big.mark = ","), " in acute need (4+)"))),
  check.names = FALSE)

# ---- styles ----
font    <- "Arial"
st_head <- createStyle(fontName = font, fontSize = 10, fontColour = "#FFFFFF", fgFill = "#404040",
                       textDecoration = "bold", halign = "center", valign = "center", wrapText = TRUE,
                       border = "TopBottomLeftRight", borderColour = "#BFBFBF")
st_body <- createStyle(fontName = font, fontSize = 10, valign = "top",
                       border = "TopBottomLeftRight", borderColour = "#BFBFBF")
st_wrap <- createStyle(wrapText = TRUE)
st_bold <- createStyle(textDecoration = "bold")
fill    <- function(col, fc = "#000000") createStyle(fgFill = col, fontColour = fc)
sev_fill <- c("1" = "#FADBD8", "2" = "#BDE3F0", "3" = "#F7BDB9", "4" = "#F08585", "4+" = "#E55353")
pop_fill <- c(Overall = "#D9D9D9", lebanese = "#DDEBF7", migrant = "#E2EFDA", prl = "#EDE2F6", prs = "#FFF2E5")
lvl_fill <- c("Household" = "#DDEBF7", "Individual" = "#FFF2CC",
              "Individual -> Household" = "#FCE4D6", "Household -> Population" = "#E2EFDA")

# safe helpers (skip when no rows / columns)
paint <- function(wb, sh, style, rows, cols) {
  if (length(rows) > 0 && length(cols) > 0)
    addStyle(wb, sh, style, rows = rows, cols = cols, gridExpand = TRUE, stack = TRUE)
}
note <- function(wb, sh, col, row, txt)
  writeComment(wb, sh, col = col, row = row,
               comment = createComment(txt, author = "MSNI", visible = FALSE, width = 3, height = 5))

# ---- result table sheet ----
add_table <- function(wb, sh, df) {
  df <- as.data.frame(lapply(df, function(x) if (is.factor(x)) as.character(x) else x), check.names = FALSE)
  addWorksheet(wb, sh)
  writeData(wb, sh, df, headerStyle = st_head)
  n <- nrow(df); k <- ncol(df); h <- names(df); rows <- seq_len(n) + 1
  freezePane(wb, sh, firstRow = TRUE)
  setRowHeights(wb, sh, rows = 1, heights = 48)
  if (n == 0) return(invisible())
  paint(wb, sh, st_body, rows, 1:k)
  first_num <- if ("n_HH" %in% h) which(h == "n_HH") else 1
  setColWidths(wb, sh, cols = 1:k, widths = ifelse(seq_len(k) < first_num, 22, 12))
  if ("Flag" %in% h) setColWidths(wb, sh, cols = which(h == "Flag"), widths = 20)
  
  # number formats
  paint(wb, sh, createStyle(numFmt = "#,##0", halign = "center"), rows,
        which(h %in% c("n_HH", "n_HH with profile", "Rank", "# sectors", "Population 2027", "# people 3+", "# people 4+")))
  pp <- which(h %in% c("# people 3+", "# people 4+"))
  paint(wb, sh, fill("#C00000", "#FFFFFF"), 1, pp)
  paint(wb, sh, st_bold, rows, pp)
  if ("Note" %in% h) setColWidths(wb, sh, cols = which(h == "Note"), widths = 40)
  paint(wb, sh, createStyle(numFmt = "0.0", halign = "center"), rows,
        which(grepl("^%", h) | grepl(" %$", h) | h %in% names(sev_fill)))
  ac <- which(grepl("^Avg", h))
  paint(wb, sh, createStyle(numFmt = "0.00", halign = "center"), rows, ac)
  
  # header colours
  for (lv in intersect(names(sev_fill), h)) paint(wb, sh, fill(sev_fill[[lv]]), 1, which(h == lv))
  need_c <- which(h %in% c("% HH in need (3+)", "% HH in severe need (4+)",
                           "% in need (3+)", "% severe need (4+)", "%", "% people 3+", "% people 4+"))
  for (cc in need_c) {
    paint(wb, sh, fill("#C00000", "#FFFFFF"), 1, cc)
    conditionalFormatting(wb, sh, cols = cc, rows = rows, type = "colourScale",
                          style = c("#FFFFFF", "#F8CBAD", "#C00000"), rule = c(0, 40, 100))
  }
  paint(wb, sh, fill("#1F4E78", "#FFFFFF"), 1, ac)
  paint(wb, sh, fill("#7F6000", "#FFFFFF"), 1, which(grepl("^% HH in need in ", h)))
  paint(wb, sh, fill("#843C0C", "#FFFFFF"), 1, which(grepl("^% HH in severe need in ", h)))
  tp <- which(grepl("^Top [1-3] ", h))
  paint(wb, sh, fill("#5B2C6F", "#FFFFFF"), 1, tp)
  paint(wb, sh, fill("#F3EAF7"), rows, tp)
  if (length(tp)) setColWidths(wb, sh, cols = which(grepl("^Top [1-3] need profile", h)), widths = 32)
  
  # population group colour on label columns
  if ("pop_group" %in% h && first_num > 1)
    for (p in names(pop_fill))
      paint(wb, sh, fill(pop_fill[[p]]), which(df$pop_group == p) + 1, 1:(first_num - 1))
  
  # low sample rows (orange)
  if ("Flag" %in% h) paint(wb, sh, fill("#F4B183"), which(!is.na(df$Flag) & df$Flag != "") + 1, 1:k)
  
  # MSNI rows in bold
  if ("Sector" %in% h) paint(wb, sh, st_bold, which(df$Sector == "MSNI") + 1, 1:k)
  if ("Rank" %in% h)   paint(wb, sh, st_bold, which(df$Rank %in% 1) + 1, 1:k)       # most common profile
  if ("Profile" %in% h) setColWidths(wb, sh, cols = which(h == "Profile"), widths = 45)
  if ("pop_group" %in% h) paint(wb, sh, st_bold, which(df$pop_group == "Overall") + 1, 1:max(1, first_num - 1))
  
  # header comments
  hc <- list(
    "n_HH"                         = "Households with a score (unweighted count). Individuals already aggregated to their household.",
    "% HH in need (3+)"            = "Weighted % of households with score 3+ = sum of weights of HH in need / sum of weights of HH with a score x 100.",
    "% in need (3+)"               = "Weighted % of households with score 3+ (in need).",
    "% HH in severe need (4+)"     = "Weighted % of households with score 4+ (severe need).",
    "% severe need (4+)"           = "Weighted % of households with score 4+. Empty for Health (max 3).",
    "Avg # sectors in need"        = "Weighted average number of sectors (out of 6) with score 3+. No sector in need = 0.",
    "Avg # sectors in severe need" = "Weighted average number of sectors (out of 5, Health excluded) with score 4+.",
    "% HH in need in 0 sectors"    = "Brown columns: weighted % of households in need in 0, 1 ... 6 sectors (row sums to 100%).",
    "4+"                           = "4+ = humind score 5. Empty = level not reachable.",
    "income_grp"                   = paste0("Household total income / weighted median of its OWN pop group (",
                                            paste(names(income_median_pop), format(income_median_pop, big.mark = ","), collapse = "; "),
                                            " LBP). Overall rows: median across all groups (", format(income_median, big.mark = ","), " LBP)."),
    "pop_group"                    = "Overall = all population groups together (grey rows); then each population group.",
    "Top 1 need profile"           = "Most common combination of sectors in need (3+), 'No sectoral need' excluded. Top % = weighted % of ALL households in the group.",
    "displacement_grp"             = "Non-displaced (incl. economic movers); Returnee (returnee IDP); IDP (incl. infra-settlement). 'Refugee' excluded.",
    "disability_grp"               = "Yes = at least one household member with a disability (hh_with_disabled).",
    "Profile"                      = "Combination of sectors in need in the same household. 'Other combinations' = all profiles outside the top 10.",
    "n_HH with profile"            = "Households with this profile (unweighted count).",
    "# sectors"                    = "Number of sectors in the profile.",
    "Population 2027"              = "2027 LRP population package (Syrians excluded).",
    "% people 3+"                  = "% of PEOPLE in households with score 3+ (households weighted by survey weight x hh_size).",
    "# people 3+"                  = "Rough estimate = Population 2027 x % people 3+. Lebanese by district, migrants by governorate, PRL/PRS by governorate (camp population only).",
    "# people 4+"                  = "Rough estimate = Population 2027 x % people 4+. Empty for Health (max 3).")
  for (nm in intersect(names(hc), h)) note(wb, sh, which(h == nm), 1, hc[[nm]])
}

# ---- text sheet (README, Algorithm, Checks) ----
add_text <- function(wb, sh, df, widths) {
  addWorksheet(wb, sh)
  writeData(wb, sh, df, headerStyle = st_head)
  rows <- seq_len(nrow(df)) + 1; k <- ncol(df)
  paint(wb, sh, st_body, rows, 1:k)
  paint(wb, sh, st_wrap, rows, 1:k)
  setColWidths(wb, sh, cols = 1:k, widths = widths)
  freezePane(wb, sh, firstRow = TRUE)
}

# ---- build workbook ----
wb <- createWorkbook()
add_text(wb, "README", readme, c(24, 120))

add_text(wb, "Algorithm", algorithm, c(15, 24, 55, 60, 45))
for (i in seq_len(nrow(algorithm))) {
  r <- i + 1; lv <- algorithm$Level[i]
  if (lv == "") {
    paint(wb, "Algorithm", fill("#595959", "#FFFFFF"), r, 1:5)
    paint(wb, "Algorithm", st_bold, r, 1:5)
  } else {
    paint(wb, "Algorithm", fill(lvl_fill[[lv]]), r, 2)
    if (algorithm$`Result (this run)`[i] != "") paint(wb, "Algorithm", fill("#FFF2CC"), r, 5)
  }
}
add_table(wb, "People_in_need", people_in_need)
add_table(wb, "People_in_need_units", pin_units)
note(wb, "Algorithm", 2, 1,
     "Blue = household; yellow = individual; orange = individuals counted per household; green = weighted % of households.")

for (nm in names(msni_tables))     add_table(wb, paste0("MSNI_", nm), msni_tables[[nm]])
for (nm in names(severity_tables)) add_table(wb, paste0("Sev_",  nm), severity_tables[[nm]])
for (nm in names(profile_tables))  add_table(wb, paste0("Profile_", nm), profile_tables[[nm]])
for (nm in names(acute_tables))    add_table(wb, paste0("Acute_Profile_", nm), acute_tables[[nm]])
add_table(wb, "Dimensions", dimensions)

add_text(wb, "Checks", checks, c(60, 18))
ck <- checks$Check
paint(wb, "Checks", fill("#E2EFDA"), grep("sanity check", ck) + 1, 1:2)                            # green = should be 0
paint(wb, "Checks", fill("#F8CBAD"), grep("received' missing|Health: HH with no score", ck) + 1, 1:2) # orange = to watch
paint(wb, "Checks", fill("#FFF2CC"), grep("LCSI|median", ck) + 1, 1:2)                             # yellow = method choice

saveWorkbook(wb, paste0("outputs/MSNI_2026_LBN_results", out_suffix, ".xlsx"), overwrite = TRUE)
cat("\nSaved: MSNI_2026_LBN_results", out_suffix, ".xlsx\n", sep = "")


# =====================================================================
# 13. MSNI ANALYSIS TABLE - one file, REACH format
#     column pop_group  = Overall / lebanese / migrant / prl / prs
#     group_var         = Overall or the framework disaggregation
#     each disaggregation: within each pop group + Overall (all groups together)
#     _IND sheets = individual level
# =====================================================================
options(survey.lonely.psu = "adjust")      # strata with 1 interview

all_pops <- c("lebanese", "migrant", "prl", "prs")

# ---- disaggregation plan (framework) ----
# v = disaggregation; p = pop groups; overall = also for all pop groups together
an_plan <- list(
  list(v = "governorate",      p = all_pops,              overall = TRUE),
  list(v = "district",         p = "lebanese",            overall = FALSE),   # Lebanese only
  list(v = "camp_name",        p = c("prl", "prs"),       overall = FALSE),   # camps only
  list(v = "hoh_gender_grp",   p = all_pops,              overall = TRUE),
  list(v = "hoh_age_grp",      p = all_pops,              overall = TRUE),
  list(v = "income_grp",       p = all_pops,              overall = TRUE, ov = "income_grp_all"),
  list(v = "displacement_grp", p = all_pops,              overall = TRUE),
  list(v = "disability_grp",   p = all_pops,              overall = TRUE))
grp_cols <- c("admin0", "pop_group", "governorate", "district", "camp_name", "hoh_gender_grp",
              "hoh_age_grp", "income_grp", "income_grp_all", "displacement_grp", "disability_grp")

# ---- household variables per sheet ----
score_vars <- c("msni_score", "comp_foodsec_score", "comp_edu_score", "comp_snfi_score",
                "comp_prot_score", "comp_wash_score", "comp_health_score")
mean_vars  <- c("n_sectors_in_need", "n_sectors_in_severe_need")   # mean + distribution
hh_sheets <- list(
  "MSNI" = c("msni_score", "msni_in_need", "msni_in_severe_need", mean_vars, paste0(mean_vars, "_dist"),
             "msni_needs_profile", "msni_acute_profile"),
  "Food Security" = c("comp_foodsec_score", "comp_foodsec_in_need", "comp_foodsec_in_severe_need",
                      "fsl_fcs_cat", "fsl_rcsi_cat", "fsl_hhs_cat_ipc", lcsi_var, "fsl_fc_phase", "fclcm_phase"),
  "Education" = c("comp_edu_score", "comp_edu_in_need", "comp_edu_in_severe_need",
                  "comp_edu_score_attendance", "comp_edu_score_disrupted"),
  "SNFI" = c("comp_snfi_score", "comp_snfi_in_need", "comp_snfi_in_severe_need",
             "comp_snfi_score_shelter_type_cat", "comp_snfi_score_shelter_issue_cat",
             "comp_snfi_score_tenure_security_cat", "comp_snfi_score_fds_cannot_cat",
             "comp_snfi_score_shelter_damage_cat"),
  "Protection" = c("comp_prot_score", "comp_prot_in_need", "comp_prot_in_severe_need",
                   "comp_prot_score_movement", "comp_prot_score_practices", "comp_prot_score_rights"),
  "WASH" = c("comp_wash_score", "comp_wash_in_need", "comp_wash_in_severe_need",
             "comp_wash_score_water_quantity", "comp_wash_score_water_quality",
             "comp_wash_score_sanitation", "comp_wash_score_hygiene"),
  "Health" = c("comp_health_score", "comp_health_in_need"))

# ---- household data: groups as text, scores 5 -> "4+", categories as text ----
msni$admin0 <- "Lebanon"
cat_vars <- setdiff(intersect(unlist(hh_sheets), names(msni)), c(score_vars, mean_vars))
msni_an <- msni |>
  mutate(across(all_of(grp_cols), as.character),
         strata_u = paste(pop_group, strata, sep = "_"),       # strata unique across pop groups
         across(any_of(mean_vars), as.character, .names = "{.col}_dist"),
         across(any_of(score_vars), ~ if_else(.x >= 5, "4+", as.character(.x))),
         across(any_of(cat_vars), as.character))

# ---- individual data: add household weight, strata and groups ----
hh_keys <- msni_an |> select(all_of(c("uuid", "weight", "strata_u", grp_cols)))
add_hh <- function(loop) {
  loop |>
    select(-any_of(c("weight", "strata", "strata_u", grp_cols))) |>
    inner_join(hh_keys, by = c("_submission__uuid" = "uuid"))
}
edu_flags <- setdiff(grep("^edu_ind_.*_d$", names(edu_ind), value = TRUE), "edu_ind_schooling_age_d")
edu_an <- add_hh(edu_ind) |>
  filter(edu_ind_age_schooling %in% 1) |>                     # school-age children (edu_age_start - edu_age_end)
  mutate(across(all_of(edu_flags), as.character))
hlt_vars <- setdiff(grep("^health_ind_healthcare_needed", names(hlt_ind), value = TRUE),
                    "health_ind_healthcare_needed")
hlt_an <- add_hh(hlt_ind) |> mutate(across(all_of(hlt_vars), as.character))
cat("Education_IND variables:", edu_flags, "\nHealth_IND variables:", hlt_vars, "\n")

# ---- run analysistools for one sheet ----
run_sheet <- function(d, vars, theme) {
  vars <- intersect(vars, names(d))
  if (nrow(d) == 0 || length(vars) == 0) return(NULL)
  design <- as_survey_design(d, strata = strata_u, weights = weight)
  make_loa <- function(g) data.frame(
    analysis_type = ifelse(vars %in% mean_vars, "mean", "prop_select_one"),
    analysis_var  = vars, group_var = g, level = 0.95)
  run_one <- function(des, g, label) tryCatch(
    suppressMessages(create_analysis(des, loa = make_loa(g), sm_separator = ".")$results_table),
    error = function(e) { cat("  !", theme, "/", label, ":", conditionMessage(e), "\n"); NULL })
  # relabel: pop_group column + group_var / group_var_value + full analysis_key
  relabel <- function(res, pop, gvar, gval_from_res = TRUE, key_from = NULL, key_to = NULL) {
    if (is.null(res) || nrow(res) == 0) return(NULL)
    res |> mutate(
      pop_group       = pop,
      group_var_value = if (gval_from_res) as.character(group_var_value) else gvar,
      group_var       = gvar,
      analysis_key    = if (is.null(key_from)) analysis_key else sub(key_from, key_to, analysis_key, fixed = TRUE))
  }
  
  out <- list()
  # 1. Overall (all pop groups together)
  r <- run_one(design, "admin0", "Overall")
  out[["overall"]] <- relabel(r, "Overall", "Overall")
  # 2. each pop group, no other breakdown
  r <- run_one(design, "pop_group", "pop_group")
  if (!is.null(r)) out[["pop"]] <- r |> mutate(pop_group = as.character(group_var_value),
                                               group_var = "Overall")
  # 3. each disaggregation: Overall + within each pop group
  for (x in an_plan) {
    v <- x$v
    if (x$overall) {
      ovar <- if (is.null(x[["ov"]])) v else x[["ov"]]   # income: Overall rows use median across all groups (exact match)
      des  <- filter(design, !is.na(.data[[ovar]]))
      out[[paste("Overall", v)]] <- relabel(run_one(des, ovar, paste("Overall", v)), "Overall", v,
                                            key_from = paste0("@/@ ", ovar, " %/% "),
                                            key_to   = paste0("@/@ ", v, " %/% "))
    }
    for (p in x$p) {
      des <- filter(design, pop_group == p, !is.na(.data[[v]]))
      if (nrow(des$variables) == 0) next
      out[[paste(p, v)]] <- relabel(run_one(des, v, paste(p, v)), p, v,
                                    key_from = paste0("@/@ ", v, " %/% "),
                                    key_to   = paste0("@/@ pop_group %/% ", p, " -/- ", v, " %/% "))
    }
  }
  bind_rows(out) |>
    filter(!(analysis_type == "prop_select_one" & is.na(analysis_var_value))) |>   # drop empty "missing" rows
    mutate(theme = theme) |>
    select(any_of(c("theme", "analysis_type", "analysis_var", "analysis_var_value", "pop_group",
                    "group_var", "group_var_value", "stat", "stat_low", "stat_upp", "n", "n_total",
                    "n_w", "n_w_total", "analysis_key")))
}

# ---- run all sheets ----
tabs <- list()
for (s in names(hh_sheets)) { cat("Running", s, "\n"); tabs[[s]] <- run_sheet(msni_an, hh_sheets[[s]], s) }
cat("Running Education_IND\n"); tabs[["Education_IND"]] <- run_sheet(edu_an, edu_flags, "Education")
cat("Running Health_IND\n");    tabs[["Health_IND"]]    <- run_sheet(hlt_an, hlt_vars,  "Health")
tabs <- tabs[c("MSNI", "Food Security", "Education", "Education_IND", "SNFI",
               "Protection", "WASH", "Health", "Health_IND")]

# ---- check: pop_group and group_var present ----
cat("\n--- check per sheet (expected: 5 pop_group values, 9 group_var values) ---\n")
for (s in names(tabs)) cat(s, ": pop_group", n_distinct(tabs[[s]]$pop_group),
                           "| group_var", n_distinct(tabs[[s]]$group_var), "\n")

# ---- People in need (unit figures from section 11b), analysis-table layout ----
pin_an <- bind_rows(lapply(c("3+", "4+"), function(k) {
  pin_units |>
    transmute(theme = Sector, analysis_type = "people_in_need_estimate",
              analysis_var = paste0(sectors[Sector], ifelse(k == "3+", "_in_need", "_in_severe_need")),
              pop_group,
              group_var       = case_when(Level == "national" ~ "Overall",
                                          grepl("governorate", Level) ~ "governorate",
                                          TRUE ~ "district"),
              group_var_value = if_else(Level == "national", "Lebanon", Unit),
              stat       = if (k == "3+") `# people 3+` else `# people 4+`,          # number of people
              stat_pct   = (if (k == "3+") `% people 3+` else `% people 4+`) / 100,  # proportion of people
              population = `Population 2027`, n = n_HH, note = Note)
})) |>
  filter(!is.na(stat)) |>                                                   # Health 4+ not applicable
  arrange(factor(theme, sec_order), analysis_var, factor(pop_group, pop_order),
          factor(group_var, c("Overall", "governorate", "district")), group_var_value)
tabs[["People_in_need"]] <- pin_an
cat("People_in_need (analysis table):", nrow(pin_an), "rows\n")

# ---- READ_ME ----
readme13 <- data.frame(
  Items = c("Population", "Rules", "Packages", "Design", "Weights", "pop_group column",
            "Disaggregation", "Needs profile", "People_in_need sheet", "Scores", "Levels", "Fields", "Run date"),
  Description = c(
    "All population groups: Lebanese, live-out migrants, PRL, PRS",
    "REACH_MSNA_2026_MSNI_final.xlsx (Lebanon MSNI rules, all sector tabs incl. Disaggregation tab)",
    paste0("humind ", packageVersion("humind"), "; analysistools ", packageVersion("analysistools")),
    "Stratified (column 'strata', unique per pop group); 95% CI. Migrant clusters not in dataset (CI slightly narrower).",
    "Survey weight 'weight' (per stratum: population share / interview share)",
    "Overall = all population groups together; otherwise lebanese / migrant / prl / prs",
    paste0("group_var = Overall (no other breakdown) or: governorate (all pop groups); district (Lebanese); ",
           "camp_name (PRL, PRS); hoh_gender_grp; hoh_age_grp (Below 18 / 18-59 / 60+); ",
           "income_grp (% of weighted median of own pop group: ",
           paste(names(income_median_pop), format(income_median_pop, big.mark = ","), collapse = ", "),
           " LBP; Overall rows: median across all groups = ", format(income_median, big.mark = ","), " LBP); ",
           "displacement_grp (Non-displaced / Returnee / IDP); disability_grp (at least one member). ",
           "Each also for pop_group = Overall, except district and camp."),
    "msni_needs_profile = combination of sectors in need (3+); msni_acute_profile = sectors at 4+ (Health excluded). All profiles listed.",
    paste0("Rough estimates of # people (stat) in need (_in_need = 3+) and acute need (_in_severe_need = 4+): 2027 LRP population ",
           "(Syrians excluded) x % of people (stat_pct; households weighted by weight x hh_size). Lebanese by district, migrants by ",
           "governorate, PRL/PRS by governorate (camp population only; gatherings excluded). Overall = sum of groups."),
    "1-4+ (4+ = humind score 5). In need = 3+, severe need = 4+",
    "Household tabs: base = households. _IND tabs: base = individuals, with household weight",
    "REACH analysis-table fields + 'pop_group' column; stat = proportion (0-1) or mean",
    as.character(Sys.Date())))

# ---- export ----
st_head13 <- createStyle(fontName = "Arial", fontColour = "#FFFFFF", fgFill = "#404040",
                         textDecoration = "bold", halign = "center", wrapText = TRUE)
wb13 <- createWorkbook()
addWorksheet(wb13, "READ_ME"); writeData(wb13, "READ_ME", readme13, headerStyle = st_head13)
setColWidths(wb13, "READ_ME", 1:2, c(18, 120))
for (s in names(tabs)) {
  addWorksheet(wb13, s); writeData(wb13, s, tabs[[s]], headerStyle = st_head13)
  freezePane(wb13, s, firstRow = TRUE); setColWidths(wb13, s, 1:15, "auto")
  cat(s, ":", nrow(tabs[[s]]), "rows\n")
}
f <- paste0("outputs/MSNI_Analysis_Table_MSNA2026", out_suffix, ".xlsx")
saveWorkbook(wb13, f, overwrite = TRUE)
cat("\nSaved:", f, "\nALL DONE\n")