WITH RECURSIVE CategoryHierarchy AS (
    SELECT catid, name, supercatid, name AS hierarchy FROM code_cat WHERE supercatid IS NULL
    UNION ALL
    SELECT c.catid, c.name, c.supercatid, ch.hierarchy || ' ⊇ ' || c.name FROM code_cat c INNER JOIN CategoryHierarchy ch ON c.supercatid = ch.catid)
SELECT ch.hierarchy AS "Axial Category", code_name.name AS "Open Code",  replace(attribute.value || ': ' || GROUP_CONCAT(code_text.seltext, ' ... '), '"', '''') AS "Code Source", cases.name AS "Theme", aspect.value AS "Aspect"
FROM code_name 
JOIN code_text ON code_name.cid = code_text.cid 
JOIN case_text ON code_text.fid = case_text.fid 
JOIN cases ON cases.caseid = case_text.caseid 
JOIN attribute ON attribute.id = code_text.fid AND attribute.name = 'Interviewee' 
JOIN CategoryHierarchy ch ON ch.catid = code_name.catid 
JOIN code_cat ON code_cat.catid = code_name.catid 
JOIN source ON source.id = code_text.fid
JOIN attribute aspect ON aspect.id = cases.caseid AND aspect.name = 'Aspect'
WHERE LOWER(code_name.name) LIKE LOWER('%SearchKeyword%') 
       OR LOWER(ch.hierarchy) LIKE LOWER('%SearchKeyword*%') 
       OR LOWER(code_text.seltext) LIKE LOWER('%SearchKeyword%')
GROUP BY ch.hierarchy, code_name.name, cases.name, aspect.value
ORDER BY ch.hierarchy;