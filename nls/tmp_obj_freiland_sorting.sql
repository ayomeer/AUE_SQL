-- View: dbu_aue_nls.tmp_obj_freiland_sorting

-- DROP VIEW dbu_aue_nls.tmp_obj_freiland_sorting;

CREATE OR REPLACE VIEW dbu_aue_nls.tmp_obj_freiland_sorting AS 
 WITH starting_point_objects AS (
         SELECT biovzgl_sf.id,
            biovzgl_sf.tmp_nr_obj_freiland_50m,
            biovzgl_sf.flaeche_m2,
            biovzgl_sf.the_geom_ptosf AS g2
           FROM dbu_aue_nls.biovzgl_sf
          WHERE biovzgl_sf.tmp_tobj_max_flaeche = 1
        )
 SELECT row_number() OVER ()::integer AS vid,
    v.id,
    v.tmp_nr_obj_freiland_50m,
    st_distance(v.the_geom_ptosf, s.g2) AS distance,
    v.tmp_tobj_max_flaeche,
    rank() OVER (PARTITION BY v.tmp_nr_obj_freiland_50m ORDER BY st_distance(v.the_geom_ptosf, s.g2))::integer AS rank_dist
   FROM dbu_aue_nls.biovzgl_sf v,
    starting_point_objects s
  WHERE v.tmp_nr_obj_freiland_50m = s.tmp_nr_obj_freiland_50m
  ORDER BY v.tmp_nr_obj_freiland_50m, st_distance(v.the_geom_ptosf, s.g2);

ALTER TABLE dbu_aue_nls.tmp_obj_freiland_sorting
  OWNER TO dbu_aue_nls_write;
GRANT ALL ON TABLE dbu_aue_nls.tmp_obj_freiland_sorting TO dbu_aue_nls_write;
COMMENT ON VIEW dbu_aue_nls.tmp_obj_freiland_sorting
  IS 'View erstellt Ranking der Freiland Teilobjekte im Bezug auf die Distanz zum Objekt mit der grössten Fläche(muss mit tmp_tobj_max_flaeche = 1 gekennzeichnet sein).';
