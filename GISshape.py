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
    import numpy as np
    import pandas as pd
    import geopandas as gpd
    from shapely.geometry import LineString
    
    def get_line(df, x1 = o_x_col, x2 = d_x_col, y1 = o_x_col, y2 = d_y_col):
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
        place = pd.concat([cvp[[o_col, o_x_col, o_y_col]].rename(columns={o_col: 'PlaceName', o_x_col:'PlaceLng', o_y_col:'PlaceLat'}),
                        cvp[[d_col, d_x_col, d_y_col]].rename(columns={d_col: 'PlaceName', d_x_col:'PlaceLng', d_y_col:'PlaceLat'})]) \
                .drop_duplicates(subset=['PlaceName']) \
                .sort_values('PlaceName') \
                .reset_index(drop=True).reset_index(names='PlaceID')
        df_copy = cvp.copy()
        df_copy = pd.merge(df_copy, place.rename(columns = {'PlaceName':o_col}), on = o_col)
        df_copy = pd.merge(df_copy, place.rename(columns = {'PlaceName':d_col}), on = d_col, suffixes=['_o','_d'])
        df_copy['OD'] = df_copy['PlaceID_o'].astype(str) + '-' + df_copy['PlaceID_d'].astype(str)
        df_copy
        # 確保小的 PlaceID 放前面，大的放後面
        df_copy['Pair'] = np.where(df_copy['PlaceID_o'] < df_copy['PlaceID_d'],
                            df_copy['PlaceID_o'].astype(str) + '-' + df_copy['PlaceID_d'].astype(str),
                            df_copy['PlaceID_d'].astype(str) + '-' + df_copy['PlaceID_o'].astype(str))
        if how != 'countd':
            countdf = df_copy.groupby(['Pair'], as_index=False).agg({count_col: 'sum'}).sort_values(count_col,ascending=False).reset_index(drop = True) # Groupby 並計算總和
            countdf[['PlaceID1', 'PlaceID2']] = countdf['Pair'].str.split('-', expand=True).astype(int) # 拆分 'Pair' 欄位為 PlaceID1 和 PlaceID2 並轉為 int 類型
            
            # 準備 place 資料表的 PlaceID 對應座標，只處理一次
            place_rename = place.rename(columns={'PlaceID': 'PlaceID1'})
            place_rename2 = place.rename(columns={'PlaceID': 'PlaceID2'})

            # 合併 PlaceID1 和 PlaceID2 的對應經緯度座標
            countdf = countdf.merge(place_rename, on='PlaceID1')\
                            .merge(place_rename2, on='PlaceID2', suffixes=['_o', '_d'])
            
            # 篩選所需欄位並重新命名
            countdf = countdf[['Pair', count_col, 'PlaceLng_o', 'PlaceLat_o', 'PlaceLng_d', 'PlaceLat_d']]\
                    .rename(columns={'PlaceLng_o': o_x_col, 
                                        'PlaceLat_o': o_y_col, 
                                        'PlaceLng_d': d_x_col, 
                                        'PlaceLat_d': d_y_col})
        elif how == 'countd':
            
            countdf = df_copy.groupby(['Pair'], as_index=False).agg({count_col: 'sum',date_col:'nunique'})

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
            countdf = countdf[['Pair', count_col, 'PlaceLng_o', 'PlaceLat_o', 'PlaceLng_d', 'PlaceLat_d']]\
                    .rename(columns={'PlaceLng_o': o_x_col, 
                                        'PlaceLat_o': o_y_col, 
                                        'PlaceLng_d': d_x_col, 
                                        'PlaceLat_d': d_y_col})
    
    countgdf = get_line(countdf)
    return countgdf