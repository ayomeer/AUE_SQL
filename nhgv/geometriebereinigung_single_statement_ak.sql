WITH geometry_dump AS (
        SELECT f.id, (st_dump(f.the_geom)).geom as the_geom_dumped
        FROM dbu_aue_nhgv.tmp_copy_lw_bew AS f),
     temp AS (
	SELECT id, st_multi(st_union(ST_CollectionExtract(st_makevalid(the_geom_dumped),3))) as the_geom_new
	FROM geometry_dump 
	GROUP BY id)   
UPDATE dbu_aue_nhgv.tmp_copy_lw_bew lw
SET the_geom = temp.the_geom_new
FROM temp
WHERE temp.id = lw.id;

