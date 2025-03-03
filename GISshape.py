import pandas as pd
import geopandas as gpd
import numpy as np
import math
from shapely.geometry import Point, LineString

def dataframe_to_point(df, lon_col, lat_col, crs="EPSG:4326", target_crs="EPSG:3826"):
    '''
    Parameters:
    df (dataframe) : 含經緯度座標欄位的dataframe
    lon_col (str) : 緯度欄位
    Lat_col (str) : 經度欄位
    crs (str) : 目前經緯度座標的座標系統，常用的為4326(WGS84)、3826(TWD97)
    target_crs：目標轉換的座標系統
    '''

    # from shapely.geometry import Point
    # import pandas as pd
    # import geopandas as gpd
    # Create Point geometries from the longitude and latitude columns
    geometry = [Point(xy) for xy in zip(df[lon_col], df[lat_col])]
    # Create a GeoDataFrame with the original CRS
    gdf = gpd.GeoDataFrame(df, geometry=geometry, crs=crs)
    # Convert the GeoDataFrame to the target CRS
    gdf = gdf.to_crs(epsg=target_crs.split(":")[1])
    return gdf

def get_line(df, x1 = 'Lon_o', x2 = 'Lon_d', y1 = 'Lat_o', y2 = 'Lat_d'):
    '''
    Parameters:
    df (dataframe) : 含經緯度座標欄位的dataframe
    x1 (str) : 起點經度欄位
    y1 (str) : 起點緯度欄位
    x2 (str) : 迄點經度欄位
    y2 (str) : 迄點緯度欄位

    預設立場：輸出為wgs84轉換的經緯度點位
    '''
    # from shapely.geometry import LineString
    # import pandas as pd
    # import geopandas as gpd
    df['geometry'] = df.apply(lambda row: LineString([(row[x1], row[y1]), (row[x2], row[y2])]), axis=1)
    gdf = gpd.GeoDataFrame(df, geometry='geometry')
    # 設定座標系統 (假設 WGS 84 / EPSG:4326)
    gdf.set_crs(epsg=4326, inplace=True)
    return gdf

def get_OD_line_shp(df, o_col, d_col, o_x_col, o_y_col, d_x_col, d_y_col, count_col, date_col, how = 'countd' ,combine = True):
    '''
    Parameters:
    df (dataframe) : 要計算的表格，例如信令資料、票證資料統計圖表
    o_col (str) : 起點的名稱 / 站序 / 可判別的欄位
    d_col (str) : 迄點的名稱 / 站序 / 可判別的欄位
    o_x_col (str) : 起點的經度 (wgs84)
    o_y_col (str) : 起點的緯度 (wgs84)
    d_x_col (str) : 迄點的經度 (wgs84)
    d_y_col (str) : 迄點的緯度 (wgs84)
    count_col (str) : 需要統計依據的量值
    combine (Boolean) : True為去定義是否要把同一對起訖端點的量合併成同一條統計量，False則僅顯示OD
    how (str) : 填入'sum'或'mean'
    OthersObject:
    countdf (dataframe):統計出來的OD表 (尚未轉成geodataframe)
    countgdf　(geodataframe):為WGS84的geodataframe
    '''

    def get_line(df, x1 = o_x_col, x2 = d_x_col, y1 = o_y_col, y2 = d_y_col):
        df['geometry'] = df.apply(lambda row: LineString([(row[x1], row[y1]), (row[x2], row[y2])]), axis=1)
        gdf = gpd.GeoDataFrame(df, geometry='geometry')
        # 設定座標系統 (假設 WGS 84 / EPSG:4326)
        gdf.set_crs(epsg=4326, inplace=True)
        return gdf

    # 1. 不合併雙向OD
    if combine == False : 
        if how != 'countd':
            countdf = df.groupby([o_col, d_col]).agg({o_x_col:'mean', o_y_col:'mean',  d_x_col:'mean', d_y_col:'mean', count_col : how}).reset_index().sort_values(count_col,ascending=False).reset_index(drop = True)
        elif how == 'countd':
            countdf = df.groupby([o_col, d_col]).agg({o_x_col:'mean', o_y_col:'mean',  d_x_col:'mean', d_y_col:'mean', count_col : 'sum', date_col:'nunique'}).reset_index()
            countdf[count_col] = countdf[count_col] / countdf[date_col]
            countdf = countdf.drop(columns = date_col).sort_values(count_col,ascending=False).reset_index(drop = True)
    else :
        place = pd.concat([df[[o_col, o_x_col, o_y_col]].rename(columns={o_col: 'PlaceName', o_x_col:'PlaceLng', o_y_col:'PlaceLat'}),
                        df[[d_col, d_x_col, d_y_col]].rename(columns={d_col: 'PlaceName', d_x_col:'PlaceLng', d_y_col:'PlaceLat'})]) \
                .drop_duplicates(subset=['PlaceName']) \
                .sort_values('PlaceName') \
                .reset_index(drop=True).reset_index(names='PlaceID')
        df_copy = df.copy()
        df_copy = pd.merge(df_copy, place.rename(columns = {'PlaceName':o_col}), on = o_col)
        df_copy = pd.merge(df_copy, place.rename(columns = {'PlaceName':d_col}), on = d_col, suffixes=['_o','_d'])
        df_copy['OD'] = df_copy['PlaceID_o'].astype(str) + '-' + df_copy['PlaceID_d'].astype(str)
        df_copy
        # 確保小的 PlaceID 放前面，大的放後面
        df_copy['Pair'] = np.where(df_copy['PlaceID_o'] < df_copy['PlaceID_d'],
                            df_copy['PlaceID_o'].astype(str) + '-' + df_copy['PlaceID_d'].astype(str),
                            df_copy['PlaceID_d'].astype(str) + '-' + df_copy['PlaceID_o'].astype(str))
        
        df_copy['PlacePair'] = np.where(df_copy['PlaceID_o'] < df_copy['PlaceID_d'],
                        df_copy[o_col].astype(str) + '-' + df_copy[d_col].astype(str),
                        df_copy[d_col].astype(str) + '-' + df_copy[o_col].astype(str))

        if how != 'countd':
            countdf = df_copy.groupby(['PlacePair','Pair'], as_index=False).agg({count_col: 'sum'}).sort_values(count_col,ascending=False).reset_index(drop = True) # Groupby 並計算總和
            countdf[['PlaceID1', 'PlaceID2']] = countdf['Pair'].str.split('-', expand=True).astype(int) # 拆分 'Pair' 欄位為 PlaceID1 和 PlaceID2 並轉為 int 類型
            
            # 準備 place 資料表的 PlaceID 對應座標，只處理一次
            place_rename = place.rename(columns={'PlaceID': 'PlaceID1'})
            place_rename2 = place.rename(columns={'PlaceID': 'PlaceID2'})

            # 合併 PlaceID1 和 PlaceID2 的對應經緯度座標
            countdf = countdf.merge(place_rename, on='PlaceID1')\
                            .merge(place_rename2, on='PlaceID2', suffixes=['_o', '_d'])
            
            # 篩選所需欄位並重新命名
            countdf = countdf[['PlacePair', count_col, 'PlaceLng_o', 'PlaceLat_o', 'PlaceLng_d', 'PlaceLat_d']]\
                    .rename(columns={'PlaceLng_o': o_x_col, 
                                        'PlaceLat_o': o_y_col, 
                                        'PlaceLng_d': d_x_col, 
                                        'PlaceLat_d': d_y_col})
        elif how == 'countd':
            
            countdf = df_copy.groupby(['Pair','PlacePair'], as_index=False).agg({count_col: 'sum',date_col:'nunique'})

            countdf[count_col] = countdf[count_col] / countdf[date_col]
            
            countdf = countdf.drop(columns = date_col).sort_values(count_col,ascending=False).reset_index(drop = True)
            # 準備 place 資料表的 PlaceID 對應座標，只處理一次
            countdf[['PlaceID1', 'PlaceID2']] = countdf['Pair'].str.split('-', expand=True).astype(int)

            place_rename = place.rename(columns={'PlaceID': 'PlaceID1'})
            place_rename2 = place.rename(columns={'PlaceID': 'PlaceID2'})
            # 合併 PlaceID1 和 PlaceID2 的對應經緯度座標
            countdf = countdf.merge(place_rename, on='PlaceID1')\
                            .merge(place_rename2, on='PlaceID2', suffixes=['_o', '_d'])
            # 篩選所需欄位並重新命名
            countdf = countdf[['PlacePair', count_col, 'PlaceLng_o', 'PlaceLat_o', 'PlaceLng_d', 'PlaceLat_d']]\
                    .rename(columns={'PlaceLng_o': o_x_col, 
                                        'PlaceLat_o': o_y_col, 
                                        'PlaceLng_d': d_x_col, 
                                        'PlaceLat_d': d_y_col})

    countgdf = get_line(countdf)
    return countgdf

def matchpolygon(polygon, pointlist , pointLat = 'PositionLat', pointLon = 'PositionLon'):
    '''
    polygon(gdf): 面狀的shp
    pointlist(df):表格，需含有經緯度資料
    pointLat(str):經度座標(WGS84)
    pointLon(str):緯度座標(WGS84)
    '''

    # import pandas as pd  #表格整理
    # import geopandas as gpd #讀取shapefile 和進行空間計算
    # from shapely.geometry import Point #計算距離

    if polygon.crs != 'EPSG:4326': #先轉回WGS
        polygon = polygon.to_crs(epsg = 4326)
    pointlist = pointlist.astype({
        pointLon: "float",
        pointLat: "float"
    })    
    geometry = [Point(xy) for xy in zip(pointlist[pointLon], pointlist[pointLat])]
    pointlist = gpd.GeoDataFrame(pointlist, geometry=geometry, crs="EPSG:4326") #把點位資料轉回'wgs84' 的經緯度座標
    pointlist = pointlist.to_crs(epsg=4326)
    pointlist_matchpolygon = gpd.sjoin(polygon, pointlist, how="right", predicate="intersects")
    pointlist_matchpolygon = pointlist_matchpolygon.drop(columns = ['geometry'])
    if 'index_right' in list(pointlist_matchpolygon.columns):
        pointlist_matchpolygon = pointlist_matchpolygon.drop(columns = 'index_right')
    if 'index_left' in list(pointlist_matchpolygon.columns):
        pointlist_matchpolygon = pointlist_matchpolygon.drop(columns = 'index_left')
    pointlist_matchpolygon = pointlist_matchpolygon[~pointlist_matchpolygon['COUNTYNAME'].isna()]
    return pointlist_matchpolygon

def get_unique_item_shp(shp, columns, folder, onlyone = True, suffix = ''):
    '''
    shp(gdf):shp，不限制型態
    columns(str) : 參照的欄位名稱
    onlyone(boolean) : 用於判斷是否要進行選擇符合條件的shp，若更改為False，則會輸出除去符合條件的shp
    suffix(str):用於當onlyone == False時的檔案名稱後綴
    '''

    # import os
    # import geopandas as gpd
    # import pandas as pd

    uniquevalue = shp[columns].unique()
    alllist = [selectitem for selectitem in uniquevalue]

    if onlyone == True:
        for selectitem in alllist :
            filename = f'{selectitem}.shp'
            outputpath = os.path.join(folder, filename)
            selectshp = shp[shp[columns] == selectitem]
            selectshp.to_file(outputpath)
    else:
        for selectitem in alllist :
            filename = f'除{selectitem}之外{suffix}.shp'
            outputpath = os.path.join(folder, filename)
            selectshp = shp[shp[columns] != selectitem]
            selectshp.to_file(outputpath)

def df_centroid(df):
    '''
    df(gdf):需要是polygon的geodataframe
    '''
    
    # 計算中心點並新增至 DataFrame
    df['centroid'] = df.geometry.centroid
    # 如果你只想看座標的經緯度，可以這樣拆開：
    df['Y'] = df['centroid'].y
    df['X'] = df['centroid'].x
    df = df.drop(columns = 'centroid')
    return df

def earth_dist(lat1, long1, lat2, long2):
    rad = math.pi / 180
    a1 = lat1 * rad
    a2 = long1 * rad
    b1 = lat2 * rad
    b2 = long2 * rad
    dis_lon = b2 - a2
    dis_lat = b1 - a1
    a = (math.sin(dis_lat / 2))**2 + math.cos(a1) * math.cos(b1) * (math.sin(dis_lon / 2))**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    R = 6378145  # 地球半徑，單位為公里
    d = R * c
    return d
