CREATE MATERIALIZED VIEW dbu_aue_gslw.tmp_1_cluster_clean AS 

	WITH t1 As ( 
	    SELECT DISTINCT 
            row_number() OVER () AS id,
            g1.id_from_zustelladresse,
            g1.id_geb,
            st_setsrid(st_makevalid(st_buffer(st_shortestline(g1.the_geom_stpointonsurface, g2.the_geom_stpointonsurface), 1::double precision)), 2056) AS the_geom_shortestline
	    FROM dbu_aue_gslw.gebaeude g1, dbu_aue_gslw.gebaeude g2
	    WHERE g1.id_from_zustelladresse = g2.id_from_zustelladresse 
        AND g1.id_geb <> g2.id_geb 
        AND st_length(st_shortestline(g1.the_geom_stpointonsurface, g2.the_geom_stpointonsurface)) <= 50::double precision 
        AND st_length(st_shortestline(g1.the_geom_stpointonsurface, g2.the_geom_stpointonsurface)) > 0::double precision
	    ORDER BY g1.id_from_zustelladresse, g1.id_geb),
	t2 As ( 
	    SELECT DISTINCT 
            row_number() OVER () AS id,
            g1.id_from_zustelladresse,
            g1.id_geb,
            st_setsrid(st_makevalid(st_buffer(st_shortestline(g1.the_geom_stpointonsurface, g2.the_geom_stpointonsurface), 1::double precision)), 2056) AS the_geom_shortestline
	    FROM dbu_aue_gslw.gebaeude g1, dbu_aue_gslw.gebaeude g2
	    WHERE g1.id_from_zustelladresse = g2.id_from_zustelladresse AND g1.id_geb <> g2.id_geb AND st_length(st_shortestline(g1.the_geom_stpointonsurface, g2.the_geom_stpointonsurface)) <= 50::double precision AND st_length(st_shortestline(g1.the_geom_stpointonsurface, g2.the_geom_stpointonsurface)) > 0::double precision
	    ORDER BY g1.id_from_zustelladresse, g1.id_geb),
	u As (
	    SELECT 
            t_buffer(st_envelope((st_dump(st_union(t1.the_geom_shortestline))).geom), 25::double precision) AS the_geom_cluster
	    FROM  t1, t2
	    WHERE intersects(t1.the_geom_shortestline, t2.the_geom_shortestline) IS TRUE),
	u2 As (
	    SELECT 
            row_number() OVER () AS nr_cluster,
            u.the_geom_cluster
	    FROM u)
    

    SELECT 
        nextval('dbu_aue_gslw.tmp_1_cluster_vid_seq'::regclass) AS vid,
        u2.nr_cluster,
        u2.the_geom_cluster,
        g.id_from_zustelladresse,
        g.id_geb
    FROM  u2, dbu_aue_gslw.gebaeude g
    WHERE st_intersects(u2.the_geom_cluster, g.the_geom) IS TRUE
    ORDER BY g.id_from_zustelladresse, u2.nr_cluster, g.id_geb
WITH DATA;

ALTER TABLE dbu_aue_gslw.tmp_1_cluster_clean
  OWNER TO dbu_aue_gslw_write;
GRANT SELECT, REFERENCES ON TABLE dbu_aue_gslw.tmp_1_cluster_clean TO dbu_aue_gslw_read;
GRANT ALL ON TABLE dbu_aue_gslw.tmp_1_cluster_clean TO dbu_aue_gslw_write;
COMMENT ON MATERIALIZED VIEW dbu_aue_gslw.tmp_1_cluster_clean
  IS 'Temporäre View fuer das Erstellen von Atlasplaenen in QGIS';
