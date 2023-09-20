SELECT seq, d.node, d.edge, cost,
e.geom As edge_geom, n.geom As node_geom, n.station
FROM
pgr_dijkstra('
SELECT gid AS id, source, target, length AS cost
FROM london_tube_lines',
(SELECT station_id
FROM london_stations WHERE station = 'Finchley Road'),
(SELECT station_id
FROM london_stations WHERE station = 'Piccadilly Circus'),
false
) AS d
LEFT JOIN london_tube_lines As e ON d.edge = e.gid
LEFT JOIN london_stations As n On d.node = n.station_id
ORDER BY d.seq;

