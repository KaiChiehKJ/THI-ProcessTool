import requests
import xml.etree.ElementTree as ET
import pandas as pd
import os 
from shapely import wkt
from shapely.errors import WKTReadingError

def create_folder(folder_name):
    """建立資料夾"""
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    return folder_name

def dataframe_to_point(df, lon_col, lat_col, crs="EPSG:4326", target_crs="EPSG:3826"):
    '''
    Parameters:
    df (dataframe) : 含經緯度座標欄位的dataframe
    lon_col (str) : 緯度欄位
    Lat_col (str) : 經度欄位
    crs (str) : 目前經緯度座標的座標系統，常用的為4326(WGS84)、3826(TWD97)
    target_crs：目標轉換的座標系統
    '''

    from shapely.geometry import Point
    import pandas as pd
    import geopandas as gpd
    # Create Point geometries from the longitude and latitude columns
    geometry = [Point(xy) for xy in zip(df[lon_col], df[lat_col])]
    # Create a GeoDataFrame with the original CRS
    gdf = gpd.GeoDataFrame(df, geometry=geometry, crs=crs)
    # Convert the GeoDataFrame to the target CRS
    gdf = gdf.to_crs(epsg=target_crs.split(":")[1])
    return gdf

def downloadfile(downloadpath, url):
    """
    下載 XML 檔案內容。
    :param downloadpath: 儲存 XML 檔案的路徑
    :param url: XML 檔案的下載 URL
    :return: XML 檔案內容
    """
    response = requests.get(url)
    if response.status_code == 200:
        with open(downloadpath, 'wb') as file:
            file.write(response.content)
        print(f"XML 檔案已下載至 {downloadpath}")
        return response.content
    else:
        raise Exception(f"無法下載檔案，HTTP 狀態碼: {response.status_code}")

def parse_vd_xml(xml_content):
    """
    解析 XML 檔案為 DataFrame，提取所有相關欄位。
    :param xml_content: XML 檔案內容
    :return: 轉換後的 DataFrame
    """
    namespace = {'ns': 'http://traffic.transportdata.tw/standard/traffic/schema/'}
    root = ET.fromstring(xml_content)

    data = []
    for vd in root.findall('.//ns:VDs/ns:VD', namespace):
        # 基本資料
        vdid = vd.find('ns:VDID', namespace).text
        sub_authority = vd.find('ns:SubAuthorityCode', namespace).text
        bi_directional = vd.find('ns:BiDirectional', namespace).text
        vd_type = vd.find('ns:VDType', namespace).text
        location_type = vd.find('ns:LocationType', namespace).text
        detection_type = vd.find('ns:DetectionType', namespace).text
        position_lon = vd.find('ns:PositionLon', namespace).text
        position_lat = vd.find('ns:PositionLat', namespace).text
        road_id = vd.find('ns:RoadID', namespace).text if vd.find('ns:RoadID', namespace) is not None else None
        road_name = vd.find('ns:RoadName', namespace).text if vd.find('ns:RoadName', namespace) is not None else None
        road_class = vd.find('ns:RoadClass', namespace).text if vd.find('ns:RoadClass', namespace) is not None else None
        location_mile = vd.find('ns:LocationMile', namespace).text if vd.find('ns:LocationMile', namespace) is not None else None

        # 道路區間
        road_section_start = vd.find('./ns:RoadSection/ns:Start', namespace).text if vd.find('./ns:RoadSection/ns:Start', namespace) is not None else None
        road_section_end = vd.find('./ns:RoadSection/ns:End', namespace).text if vd.find('./ns:RoadSection/ns:End', namespace) is not None else None

        # DetectionLinks 資料
        for link in vd.findall('./ns:DetectionLinks/ns:DetectionLink', namespace):
            link_id = link.find('ns:LinkID', namespace).text
            bearing = link.find('ns:Bearing', namespace).text if link.find('ns:Bearing', namespace) is not None else None
            road_direction = link.find('ns:RoadDirection', namespace).text if link.find('ns:RoadDirection', namespace) is not None else None
            lane_num = link.find('ns:LaneNum', namespace).text if link.find('ns:LaneNum', namespace) is not None else None
            actual_lane_num = link.find('ns:ActualLaneNum', namespace).text if link.find('ns:ActualLaneNum', namespace) is not None else None

            # 將資料加入清單
            data.append({
                "VDID": vdid,
                "SubAuthorityCode": sub_authority,
                "BiDirectional": bi_directional,
                "VDType": vd_type,
                "LocationType": location_type,
                "DetectionType": detection_type,
                "PositionLon": position_lon,
                "PositionLat": position_lat,
                "RoadID": road_id,
                "RoadName": road_name,
                "RoadClass": road_class,
                "LocationMile": location_mile,
                "RoadSectionStart": road_section_start,
                "RoadSectionEnd": road_section_end,
                "LinkID": link_id,
                "Bearing": bearing,
                "RoadDirection": road_direction,
                "LaneNum": lane_num,
                "ActualLaneNum": actual_lane_num,
            })

    return pd.DataFrame(data)

def download_and_parce_VD(downloadfolder, url):
    download_path = os.path.join(downloadfolder, 'VD.xml')
    csv_output_path = os.path.join(downloadfolder, 'VD.csv')
    
    xml_content = downloadfile(download_path, url = url)
    # 解析 XML
    df = parse_vd_xml(xml_content)
    # 儲存為 CSV
    df.to_csv(csv_output_path, index=False, encoding="big5")
    return df
    
def parse_ETag_xml(xml_content):
    """
    解析 ETag XML 檔案為 DataFrame。
    :param xml_content: XML 檔案內容
    :return: 轉換後的 DataFrame
    """
    namespace = {'ns': 'http://traffic.transportdata.tw/standard/traffic/schema/'}
    root = ET.fromstring(xml_content)

    data = []
    for etag in root.findall('.//ns:ETags/ns:ETag', namespace):
        # 提取 ETag 資訊
        gantry_id = etag.find('ns:ETagGantryID', namespace).text
        link_id = etag.find('ns:LinkID', namespace).text
        location_type = etag.find('ns:LocationType', namespace).text
        position_lon = etag.find('ns:PositionLon', namespace).text
        position_lat = etag.find('ns:PositionLat', namespace).text
        road_id = etag.find('ns:RoadID', namespace).text if etag.find('ns:RoadID', namespace) is not None else None
        road_name = etag.find('ns:RoadName', namespace).text if etag.find('ns:RoadName', namespace) is not None else None
        road_class = etag.find('ns:RoadClass', namespace).text if etag.find('ns:RoadClass', namespace) is not None else None
        road_direction = etag.find('ns:RoadDirection', namespace).text if etag.find('ns:RoadDirection', namespace) is not None else None
        location_mile = etag.find('ns:LocationMile', namespace).text if etag.find('ns:LocationMile', namespace) is not None else None

        # 道路區間
        road_section_start = etag.find('./ns:RoadSection/ns:Start', namespace).text if etag.find('./ns:RoadSection/ns:Start', namespace) is not None else None
        road_section_end = etag.find('./ns:RoadSection/ns:End', namespace).text if etag.find('./ns:RoadSection/ns:End', namespace) is not None else None

        # 加入清單
        data.append({
            "ETagGantryID": gantry_id,
            "LinkID": link_id,
            "LocationType": location_type,
            "PositionLon": position_lon,
            "PositionLat": position_lat,
            "RoadID": road_id,
            "RoadName": road_name,
            "RoadClass": road_class,
            "RoadDirection": road_direction,
            "LocationMile": location_mile,
            "RoadSectionStart": road_section_start,
            "RoadSectionEnd": road_section_end,
        })

    return pd.DataFrame(data)

def download_and_parce_ETag(downloadfolder, url):
    download_path = os.path.join(downloadfolder, 'ETag.xml')
    csv_output_path = os.path.join(downloadfolder, 'Etag.csv')

    ETag = downloadfile(downloadpath = download_path, url = urls['ETag'])
    ETag = parse_ETag_xml(ETag)
    ETag.to_csv(csv_output_path, index='False', encoding="big5")
    return ETag

def parse_CCTV_xml(xml_content):
    # Parse the XML content
    root = ET.fromstring(xml_content)
    
    # Define the namespace used in the XML
    namespace = {'ns': 'http://traffic.transportdata.tw/standard/traffic/schema/'}

    # Extract data for each CCTV element
    cctv_data = []
    for cctv in root.findall(".//ns:CCTV", namespace):
        cctv_id = cctv.find("ns:CCTVID", namespace).text if cctv.find("ns:CCTVID", namespace) is not None else None
        sub_authority_code = cctv.find("ns:SubAuthorityCode", namespace).text if cctv.find("ns:SubAuthorityCode", namespace) is not None else None
        link_id = cctv.find("ns:LinkID", namespace).text if cctv.find("ns:LinkID", namespace) is not None else None
        video_stream_url = cctv.find("ns:VideoStreamURL", namespace).text if cctv.find("ns:VideoStreamURL", namespace) is not None else None
        location_type = cctv.find("ns:LocationType", namespace).text if cctv.find("ns:LocationType", namespace) is not None else None
        position_lon = cctv.find("ns:PositionLon", namespace).text if cctv.find("ns:PositionLon", namespace) is not None else None
        position_lat = cctv.find("ns:PositionLat", namespace).text if cctv.find("ns:PositionLat", namespace) is not None else None
        road_id = cctv.find("ns:RoadID", namespace).text if cctv.find("ns:RoadID", namespace) is not None else None
        road_name = cctv.find("ns:RoadName", namespace).text if cctv.find("ns:RoadName", namespace) is not None else None
        road_class = cctv.find("ns:RoadClass", namespace).text if cctv.find("ns:RoadClass", namespace) is not None else None
        road_direction = cctv.find("ns:RoadDirection", namespace).text if cctv.find("ns:RoadDirection", namespace) is not None else None
        start_section = cctv.find("ns:RoadSection/ns:Start", namespace).text if cctv.find("ns:RoadSection/ns:Start", namespace) is not None else None
        end_section = cctv.find("ns:RoadSection/ns:End", namespace).text if cctv.find("ns:RoadSection/ns:End", namespace) is not None else None
        location_mile = cctv.find("ns:LocationMile", namespace).text if cctv.find("ns:LocationMile", namespace) is not None else None

        # Append the extracted data as a dictionary
        cctv_data.append({
            "CCTVID": cctv_id,
            "SubAuthorityCode": sub_authority_code,
            "LinkID": link_id,
            "VideoStreamURL": video_stream_url,
            "LocationType": location_type,
            "PositionLon": position_lon,
            "PositionLat": position_lat,
            "RoadID": road_id,
            "RoadName": road_name,
            "RoadClass": road_class,
            "RoadDirection": road_direction,
            "StartSection": start_section,
            "EndSection": end_section,
            "LocationMile": location_mile
        })

    # Convert the list of dictionaries to a Pandas DataFrame
    df = pd.DataFrame(cctv_data)
    return df

def download_and_parce_CCTV(downloadfolder, url):
    download_path = os.path.join(downloadfolder,'CCTV.xml')
    csv_output_path = os.path.join(downloadfolder, 'CCTV.csv')
    CCTV = downloadfile(download_path, urls['CCTV'])
    CCTV = parse_CCTV_xml(CCTV)
    CCTV.to_csv(csv_output_path, index = False, encoding='big5')
    return CCTV

def parse_ETagPair_xml(xml_content):
    # Parse the XML content
    root = ET.fromstring(xml_content)
    
    # Define the namespace
    ns = {'ns': 'http://traffic.transportdata.tw/standard/traffic/schema/'}
    
    # Extract ETagPairs
    e_tag_pairs = []
    for pair in root.findall('.//ns:ETagPair', ns):
        e_tag_pairs.append({
            'ETagPairID': pair.find('ns:ETagPairID', ns).text,
            'StartETagGantryID': pair.find('ns:StartETagGantryID', ns).text,
            'EndETagGantryID': pair.find('ns:EndETagGantryID', ns).text,
            'Description': pair.find('ns:Description', ns).text,
            'Distance': float(pair.find('ns:Distance', ns).text),
            'StartLinkID': pair.find('ns:StartLinkID', ns).text,
            'EndLinkID': pair.find('ns:EndLinkID', ns).text,
            'Geometry': pair.find('ns:Geometry', ns).text,
        })
    
    # Convert to DataFrame
    df = pd.DataFrame(e_tag_pairs)
    return df

def safe_wkt_loads(wkt_string):
    try:
        return wkt.loads(wkt_string)
    except (WKTReadingError, AttributeError):  # 捕獲解析錯誤
        return None  # 返回空幾何物件

def main():
    urls = {"VD" : "https://tisvcloud.freeway.gov.tw/history/motc20/VD.xml",
            "ETagPair" : "https://tisvcloud.freeway.gov.tw/history/motc20/ETagPair.xml",
            "ETag" : "https://tisvcloud.freeway.gov.tw/history/motc20/ETag.xml",
            "CCTV" : "https://tisvcloud.freeway.gov.tw/history/motc20/CCTV.xml"}


    downloadfolder = create_folder(os.path.join(os.getcwd(),'高公局靜態資料清單'))
    shpfolder = create_folder(os.path.join(downloadfolder, 'shp'))

    '''VD下載'''
    VD = download_and_parce_VD(downloadfolder=downloadfolder, url = urls['VD']) 
    VD = VD.rename(columns = {'SubAuthorityCode':'AuthCode',
                            'BiDirectional':'BiDirect',
                            'LocationType':'LocType',
                            'DetectionType':'DetectType',
                            'PositionLon':'Lon',
                            'PositionLat':'Lat',
                            'RoadSectionStart':'RdStart',
                            'RoadSectionEnd':'RdEnd',
                            'RoadDirection':'Direction' ,
                            'ActualLaneNum':'ActLaneNum',
                            'LocationMile':'LocMile'})# 考量shp的欄位有10字母限制
    VD = dataframe_to_point(VD, lon_col='Lon', lat_col='Lat')
    VD.to_file(os.path.join(shpfolder,'VD.shp'))
    del VD


    '''Etag下載'''
    ETag = download_and_parce_ETag(downloadfolder=downloadfolder, url = urls['ETag'])
    ETag = ETag.rename(columns={'ETagGantryID':'GantryID',
                                'LocationType':'LocType',
                                'PositionLon':'Lon',
                                'PositionLat':'Lat',
                                'RoadDirection':'Direction',
                                'LocationMile':'LocMile',
                                'RoadSectionStart':'RdStart',
                                'RoadSectionEnd':'RdEnd'})# 考量shp的欄位有10字母限制
    ETag = dataframe_to_point(ETag, lon_col='Lon', lat_col='Lat')
    ETag.to_file(os.path.join(shpfolder,'Etag.shp'))
    del ETag

    '''CCTV下載'''
    CCTV = download_and_parce_CCTV(downloadfolder=downloadfolder, url=urls['CCTV'])
    CCTV = CCTV.rename(columns={'SubAuthorityCode':'AuthCode',
                                'VideoStreamURL':'StreamURL',
                                'LocationType':'LocType',
                                'PositionLon':'Lon',
                                'PositionLat':'Lat',
                                'RoadDirection':'Direction',
                                'StartSection':'RdStart',
                                'EndSection':'RdEnd',
                                'LocationMile':'LocMile'})
    CCTV = dataframe_to_point(CCTV, lon_col='Lon', lat_col='Lat')
    CCTV.to_file(os.path.join(shpfolder,'CCTV.shp'))

    '''ETagPair 目前不能轉為gdf'''
    download_path = os.path.join(downloadfolder,'ETagPair.xml')
    csv_output_path = os.path.join(downloadfolder, 'ETagPair.csv')
    ETagPair = downloadfile(download_path, urls['ETagPair'])
    ETagPair = parse_ETagPair_xml(ETagPair)
    ETagPair = ETagPair.rename(columns = {'StartETagGantryID':'StartID',
                                        'EndETagGantryID':'EndID',
                                        'Description':'Info',
                                        'StartLinkID':'StartLink',
                                        'EndLinkID':'EndLink',
                                        'Geometry':'geometry'})

if __name__ == '__main__':
    main()