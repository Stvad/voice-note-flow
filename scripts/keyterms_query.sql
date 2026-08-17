-- Extract alias usage from the Knowledge Medium graph, as input to
-- gen_keyterms.py. Bind ?1 to the workspace id.
--
--   npx @knowledge-medium/agent-cli sql all \
--     "$(sed '/^--/d;/^$/d' scripts/keyterms_query.sql)" \
--     '["<workspace-id>"]' -p <profile> > aliases.json
--
-- The sed is required, not cosmetic: the CLI's argument parser reads a leading
-- "--" comment line as an unknown option and refuses to run.
--
-- Then convert to CSV (alias, block_id, usage_count are the columns
-- gen_keyterms.py actually reads) and run the generator.
--
-- The three exclusions below are what separate vocabulary from plumbing;
-- without them the top of the list is SRS scheduling metadata, not language:
--
--   1. srs-sm2.5 sources. Spaced-repetition cards write their scheduling as
--      body text — literally "[[factor]]:2.5" — so `interval` and `factor`
--      outrank every real term by an order of magnitude. Note that filtering
--      on source_field does NOT catch this: these are genuine body-text
--      wikilinks, not property projections.
--   2. daily-note targets. Date references dwarf everything and the
--      post-processing prompt already tags dates via its own rule.
--   3. Bracket noise and bare UUIDs, which are reference machinery rather
--      than anything a person says out loud.

SELECT
  r.alias                     AS alias,
  r.target_id                 AS block_id,
  COUNT(*)                    AS usage_count,
  COUNT(DISTINCT r.source_id) AS distinct_sources,
  b.content                   AS content_preview,
  (SELECT GROUP_CONCAT(t.type)
     FROM block_types t
    WHERE t.block_id = r.target_id) AS block_types
FROM block_references r
JOIN blocks b
  ON b.id = r.target_id
 AND b.deleted = 0
WHERE r.workspace_id = ?1
  AND NOT EXISTS (
    SELECT 1 FROM block_types st
     WHERE st.block_id = r.source_id
       AND st.type IN ('srs-sm2.5', 'matrix-message', 'readwise-highlight',
                       'readwise-document', 'readwise-note', 'block-type',
                       'extension'))
  AND NOT EXISTS (
    SELECT 1 FROM block_types dt
     WHERE dt.block_id = r.target_id
       AND dt.type = 'daily-note')
  AND r.alias NOT LIKE '%[[%'
  AND r.alias NOT GLOB '[0-9a-f]*-[0-9a-f]*-[0-9a-f]*-[0-9a-f]*-[0-9a-f]*'
GROUP BY r.alias, r.target_id
HAVING COUNT(*) >= 4
ORDER BY usage_count DESC;
