-- Unify localized suggestions: Germany (hand-curated) + FR/ES/MX/BR/IT/NL/PL/TR (auto gap)
CREATE OR REPLACE TABLE localized_suggestions AS
SELECT product, country, theme, keyword, tier, sv FROM localized_de
UNION ALL
SELECT product, country, theme, keyword, tier, sv FROM localized_loc
