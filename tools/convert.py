#!/usr/bin/env python3
"""Convert the official ZIP without simplifying, repairing, or guessing geometry."""
import argparse, collections, csv, datetime, hashlib, json, pathlib, tempfile, zipfile
import shapefile, openpyxl
from pyproj import CRS, Transformer
from shapely.geometry import shape, mapping
from shapely.ops import transform
from shapely.geometry.polygon import orient
from shapely.validation import explain_validity
ROOT=pathlib.Path(__file__).resolve().parents[1]
SOURCE='https://data.seoul.go.kr/dataList/OA-22712/F/1/datasetView.do'
DISTRICTS=dict(zip('11110 11140 11170 11200 11215 11230 11260 11290 11305 11320 11350 11380 11410 11440 11470 11500 11530 11545 11560 11590 11620 11650 11680 11710 11740'.split(),'종로구 중구 용산구 성동구 광진구 동대문구 중랑구 성북구 강북구 도봉구 노원구 은평구 서대문구 마포구 양천구 강서구 구로구 금천구 영등포구 동작구 관악구 서초구 강남구 송파구 강동구'.split()))
CATEGORIES={'BZ101':'신속통합기획','BZ102':'재개발','BZ103':'재개발','BZ104':'재건축','BZ105':'재건축','BZ201':'모아타운','BZ204':'재건축','BZ205':'재개발'}
p=argparse.ArgumentParser();p.add_argument('zip',nargs='?',default=str(ROOT/'original/seoulplan-202609.zip'));p.add_argument('--release-date',default='2026-09-10');p.add_argument('--source-filename',default='532_UQ120_도시계획사업(서울플랜+)_202609.zip');a=p.parse_args()
with tempfile.TemporaryDirectory() as td:
 d=pathlib.Path(td)
 with zipfile.ZipFile(a.zip) as z:
  for n in z.namelist():
   if not n.endswith('/'): (d/n.split('/')[-1]).write_bytes(z.read(n))
 w=openpyxl.load_workbook(next(d.glob('*.xlsx')),data_only=True)
 types={r[4]:r[3] for r in w.worksheets[0].iter_rows(min_row=2,values_only=True) if r[4]}
 stages={r[2]:r[3] for r in w.worksheets[1].iter_rows(min_row=2,values_only=True) if r[2]}
 shp=next(d.glob('*.shp'));crs=CRS.from_wkt(shp.with_suffix('.prj').read_text())
 tx=Transformer.from_crs(crs,4326,always_xy=True,allow_ballpark=False)
 r=shapefile.Reader(str(shp),encoding='cp949'); features=[]; excluded=[];related=collections.defaultdict(list)
 for rec in r.records():
  if rec['SCLAS_CL']=='BZ402': related[rec['DGM_NM']].append({'name':rec['DGM_NM'],'type':types.get(rec['SCLAS_CL']),'stage':stages.get(rec['PROPEL_CD']),'stage_code':rec['PROPEL_CD'],'source_id':rec['PRESENT_SN']})
 for sr in r.iterShapeRecords():
  v=sr.record.as_dict();code=v['SCLAS_CL']
  if code not in CATEGORIES: continue
  g=shape(sr.shape.__geo_interface__)
  if g.is_empty or g.geom_type not in ('Polygon','MultiPolygon') or not g.is_valid:
   excluded.append({'id':v['PRESENT_SN'],'name':v['DGM_NM'],'reason':explain_validity(g)});continue
  g=transform(tx.transform,g)
  if g.geom_type=='Polygon':g=orient(g,sign=1)
  else:
   from shapely.geometry import MultiPolygon
   g=MultiPolygon([orient(part,sign=1) for part in g.geoms])
  assert g.is_valid and 126<g.bounds[0]<128 and 37<g.bounds[1]<38.5
  anchor=g.representative_point(); props={'name':v['DGM_NM'],'district':DISTRICTS.get(v['SIGNGU_SE'],'자료 없음'),'district_code':v['SIGNGU_SE'],'type':CATEGORIES[code],'subtype':types.get(code,code),'type_code':code,'stage':stages.get(v['PROPEL_CD'],'자료 없음'),'stage_code':v['PROPEL_CD'],'households':None,'households_note':'서울플랜+ 원본에 세대수 항목 없음','source_area_m2':v['DGM_AR'] if v['DGM_AR']>0 else None,'record_date':str(v['CREATE_DAT']) if v['CREATE_DAT'] else None,'source_release':a.release_date,'source_url':SOURCE,'source_id':v['PRESENT_SN'],'group':v['GRP'] or None,'anchor':[anchor.x,anchor.y],'related_records':related[v['DGM_NM']],'raw':{k:str(val) if isinstance(val,datetime.date) else val for k,val in v.items()}}
  features.append({'type':'Feature','id':v['PRESENT_SN'],'properties':props,'geometry':mapping(g),'bbox':list(g.bounds)})
 features.sort(key=lambda f:(f['properties']['district'],f['properties']['name'],f['id']))
 out=ROOT/'data';out.mkdir(exist_ok=True)
 def save(n,o): (out/n).write_text(json.dumps(o,ensure_ascii=False,separators=(',',':')),encoding='utf-8')
 save('projects.geojson',{'type':'FeatureCollection','features':features})
 with (out/'coverage.csv').open('w',encoding='utf-8-sig',newline='') as file:
  writer=csv.writer(file);writer.writerow(['원본ID','구역명','자치구','유형','세부유형','단계','공개본'])
  for f in features:
   v=f['properties'];writer.writerow([v[k] for k in ['source_id','name','district','type','subtype','stage','source_release']])
 counts=dict(collections.Counter(f['properties']['type'] for f in features))
 meta={'title':'서울시 도시계획사업 현황(서울플랜+) 공간정보','source_url':SOURCE,'source_file':a.source_filename,'release_date':a.release_date,'retrieved_date':datetime.date.today().isoformat(),'license':'공공누리 제4유형: 출처표시·상업적 이용금지·변경금지','notice':'서울시 참고용 사업 경계이며 법적 효력이 없습니다. 신속통합기획·모아타운은 법정 정비구역 지정과 다를 수 있습니다.','source_crs':crs.to_string(),'source_wkt':crs.to_wkt(),'target_crs':'EPSG:4326 (longitude, latitude)','transform':tx.description,'sha256':hashlib.sha256(pathlib.Path(a.zip).read_bytes()).hexdigest(),'source_records':len(r),'feature_count':len(features),'counts':counts,'excluded':excluded,'types':types,'stages':stages}
 save('metadata.json',meta)
 print(json.dumps({'count':len(features),'categories':counts,'excluded':excluded,'transform':tx.description},ensure_ascii=False,indent=2))
