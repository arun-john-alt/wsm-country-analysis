-- Cached intent classification (LLM-reasoned). keep=false -> excluded from suggestions.
-- Question/how-to PROBLEM queries that map to a product capability are KEPT (not listed here).
-- Only intent-junk is listed: consumer, DIY-free, off-product, competitor-info, definitional.
CREATE OR REPLACE TABLE keyword_intent AS
SELECT * FROM UNNEST([
  STRUCT('what active directory' AS keyword_text, false AS keep, 'definitional' AS reason),
  ('how to reset password on windows 8', false, 'consumer'),
  ('how to reset windows login password', false, 'consumer'),
  ('how to find disabled users in active directory using powershell', false, 'diy-free'),
  ('how to track employee productivity', false, 'off-product'),
  ('what is netwrix auditor', false, 'competitor-info'),
  ('how to get 2fa', false, 'consumer-vague')
]);
