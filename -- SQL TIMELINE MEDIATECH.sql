-- SQL: TIMELINE MEDIATECH
WITH RECURSIVE CategoryHierarchy AS (
    SELECT 
        catid,
        name,
        supercatid,
        name AS hierarchy
    FROM 
        code_cat
    WHERE 
        supercatid IS NULL

    UNION ALL

    SELECT 
        c.catid,
        c.name,
        c.supercatid,
        c.name || ' ⊆ ' || ch.hierarchy
    FROM 
        code_cat c
    INNER JOIN CategoryHierarchy ch ON c.supercatid = ch.catid
)
SELECT 
    SUBSTR(code_name.memo, 
           INSTR(code_name.memo, '|Timeline： ') + LENGTH('|Timeline： '), 
           INSTR(SUBSTR(code_name.memo, INSTR(code_name.memo, '|Timeline： ') + LENGTH('|Timeline： ')), '|') - 1) AS "Timeline",
    code_name.name AS "Code Name",
    MIN(CH.hierarchy) AS "Category",
    GROUP_CONCAT(
        DISTINCT source.name || ' (' || code_in_file_count.count || ')'
    ) AS "File Freq",
    COUNT(DISTINCT source.id) AS "Count",
    COALESCE(code_name.memo, '') AS "Code Memo"
FROM 
    source
LEFT JOIN 
    code_text ON code_text.fid = source.id
LEFT JOIN 
    code_name ON code_name.cid = code_text.cid
LEFT JOIN 
    CategoryHierarchy CH ON CH.catid = code_name.catid
LEFT JOIN (
    SELECT 
        code_text.fid,
        code_text.cid,
        COUNT(*) AS count
    FROM 
        code_text
    GROUP BY 
        code_text.fid, code_text.cid
) AS code_in_file_count ON code_in_file_count.fid = source.id AND code_in_file_count.cid = code_name.cid
WHERE 
    code_name.memo IS NOT NULL AND code_name.memo <> ''
GROUP BY 
    "Timeline", code_name.name, code_name.catid -- Grouping by Timeline, Code Name, and catid
ORDER BY 
    "Timeline" DESC, "Code Name", "Category", "File Freq", "Count", "Code Memo"; -- Sorting by Timeline in descending order

