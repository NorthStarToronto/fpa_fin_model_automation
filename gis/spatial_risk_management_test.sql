SELECT 
	ln_l.project, 
	ln_l.type,
	count(pt_r.id) as nest_risk_num
FROM linear_projects ln_l
LEFT JOIN raptor_nests pt_r
ON ST_DWithin(ln_l.geom::geography, pt_r.geom::geography, 482.5)
GROUP BY ln_l.project, ln_l.type
ORDER BY nest_risk_num DESC;