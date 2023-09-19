SELECT 
	nest_id, ST_Buffer(ST_Transform(geom,26913), 200) as geom , 200 as buf_dist 
FROM baea_nests

UNION ALL 

SELECT 
	nest_id, ST_Buffer(ST_Transform(geom,26913), 400) as geom , 200 as buf_dist 
FROM baea_nests

UNION ALL

SELECT 
	nest_id, ST_Buffer(ST_Transform(geom,26913), 600) as geom , 200 as buf_dist 
FROM baea_nests