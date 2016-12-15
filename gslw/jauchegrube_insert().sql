-- Function: dbu_aue_gslw.jauchegrube_insert()

-- DROP FUNCTION dbu_aue_gslw.jauchegrube_insert();

CREATE OR REPLACE FUNCTION dbu_aue_gslw.jauchegrube_insert()
  RETURNS trigger AS
$BODY$

DECLARE
tmpid integer;
tmpeg character varying;
tmpbfs character varying;
tmpau character varying;
tmpao character varying;
tmps1 character varying;
tmps2 character varying;
tmps3 character varying;
-- keine Konstanten oder Variablendeklarationen für die Funktion

begin
-- * AV-Daten abfragen

-- Geometriedaten überprüfen und neu rechnen
NEW.the_geom := ST_SetSRID(NEW.the_geom,2056);
NEW.x_koordinate := st_x(NEW.the_geom);
NEW.y_koordinate := st_y(NEW.the_geom);

-- Datumswerte von Eingabe korrigieren

IF to_char(NEW.dpbeginn, 'YYYY') = '1888' then NEW.dpbeginn = null;END IF;
IF to_char(NEW.letztekontr, 'YYYY') = '1888' then NEW.letztekontr = null;END IF;
IF to_char(NEW.dat_standortanfrage, 'YYYY') = '1888' then NEW.dat_standortanfrage = null;END IF;
IF to_char(NEW.kontr_status_frist, 'YYYY') = '1888' then NEW.kontr_status_frist = null;END IF;

-- Grundstückdaten etc. aus AV abfragen

Select g.egris_egrid into tmpeg from dbu_tb_mop.grundstueck_resf_pub as g WHERE st_within(NEW.the_geom,g.the_geom) is true FETCH FIRST ROW ONLY;
NEW.ext_egrid := tmpeg;
NEW.ext_grundstck := g.nummer from dbu_tb_mop.grundstueck_resf_pub as g WHERE g.egris_egrid = tmpeg;
NEW.ext_ortschaft := g.ortschaftsname from dbu_tb_mop.grundstueck_resf_pub as g WHERE g.egris_egrid = tmpeg; 
Select m.bfsnr into tmpbfs from dbu_tb_mop.mbsf as m WHERE st_within(NEW.the_geom,m.geometrie) is true FETCH FIRST ROW ONLY;
NEW.ext_gemeinde := tmpbfs; 

--  * Alpnamen aus LW-Daten aktualisieren
NEW.ext_alpname = a."Alpname" from pub_lw_ln.alpnamen as a 
 where st_within (NEW.the_geom, a.the_geom) is true;

-- * Flurname aus AV Daten holen
IF NEW.flurbez is null then
	NEW.flurbez := f.name from dbu_tb_av.nomenklatur_flurname as f
	where st_within (NEW.the_geom, f.geometrie) is true
	FETCH FIRST ROW ONLY;
END IF;

-- * Daten aus zugeordnetem Gebäude abfragen
IF NEW.id_from_gebaeude is not null then
	NEW.ext_lbnr := g.ext_lbnr from dbu_aue_gslw.gebaeude as g where NEW.id_from_gebaeude = g.id_geb;
	NEW.ext_av_typ_gebaeude := g.ext_av_typ_gebaeude from dbu_aue_gslw.gebaeude as g where NEW.id_from_gebaeude = g.id_geb;
	NEW.help_kt_id := g.help_bew_id_gebaeude from dbu_aue_gslw.gebaeude as g where NEW.id_from_gebaeude = g.id_geb;
	NEW.ablagenr := g.ablagenr from dbu_aue_gslw.gebaeude as g where NEW.id_from_gebaeude = g.id_geb;
END IF;

-- * Gewässerschutzdaten abfragen

NEW.ext_gschareal := 'Areal' from pub_ue_gs.areal as gs 
    WHERE st_within(NEW.the_geom,gs.the_geom) is true FETCH FIRST ROW ONLY;

Select 'S1' into tmps1 from pub_ue_gs.gszonen as gsz where st_within(NEW.the_geom,gsz.the_geom) is true and left(gsz."Objektart",2) = 'S1' FETCH FIRST ROW ONLY;
Select 'S2' into tmps2 from pub_ue_gs.gszonen as gsz where st_within(NEW.the_geom,gsz.the_geom) is true and left(gsz."Objektart",2) = 'S2' FETCH FIRST ROW ONLY;
Select 'S3' into tmps3 from pub_ue_gs.gszonen as gsz where st_within(NEW.the_geom,gsz.the_geom) is true and left(gsz."Objektart",2) = 'S3' FETCH FIRST ROW ONLY;

Select 'Ao' into tmpao from pub_ue_gs.bereich_ao as ao where st_within(NEW.the_geom,ao.the_geom) is true FETCH FIRST ROW ONLY;
Select 'Au' into tmpau from pub_ue_gs.bereich_au as au where st_within(NEW.the_geom,au.the_geom) is true FETCH FIRST ROW ONLY;

IF tmps3 = 'S3' then NEW.ext_gschzone := 'S3';END IF;
IF tmps2 = 'S2' then NEW.ext_gschzone := 'S2';END IF;
IF tmps1 = 'S1' then NEW.ext_gschzone := 'S1';END IF;

IF tmpau = 'Au' then NEW.ext_gschbereich := 'Au'; END IF;
IF tmpao = 'Ao' then NEW.ext_gschbereich := 'Ao'; END IF;
IF tmpau = 'Au' and tmpao = 'Ao' then NEW.ext_gschbereich := 'Ao und Au'; END IF;


-- Datum setzen
NEW.created := LOCALTIMESTAMP(0);
NEW.lastmodified := LOCALTIMESTAMP(0);
-- Rückgabe der Variablen zur Speicherung
RETURN NEW;
END;



$BODY$
  LANGUAGE plpgsql VOLATILE
  COST 100;
ALTER FUNCTION dbu_aue_gslw.jauchegrube_insert()
  OWNER TO postgres;
GRANT EXECUTE ON FUNCTION dbu_aue_gslw.jauchegrube_insert() TO public;
GRANT EXECUTE ON FUNCTION dbu_aue_gslw.jauchegrube_insert() TO postgres;
GRANT EXECUTE ON FUNCTION dbu_aue_gslw.jauchegrube_insert() TO dbu_aue_gslw_write;
