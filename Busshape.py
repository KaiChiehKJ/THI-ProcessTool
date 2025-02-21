import geopandas as gpd
import pandas as pd
from shapely.geometry import LineString, Point
from shapely.ops import substring

def snap_points_to_line(stops_gdf, routes_gdf, 
                        route_routename_col, route_direction_col, 
                        seq_routename_col, seq_direction_col, 
                        seq_lat_col, seq_lng_col):
    """
    將公車站點 (stops_gdf) 投影到公車路線 (routes_gdf) 上，並動態帶入欄位名稱。
    Parameters:
        stops_gdf (GeoDataFrame): 包含公車站點的 GeoDataFrame。
        routes_gdf (GeoDataFrame): 包含公車路線的 GeoDataFrame。
        route_routename_col (str): 路線名稱欄位名稱。
        route_direction_col (str): 路線方向欄位名稱。
        seq_routename_col (str): 站點路線名稱欄位名稱。
        seq_direction_col (str): 站點方向欄位名稱。
        seq_lat_col (str): 站點緯度欄位名稱。
        seq_lng_col (str): 站點經度欄位名稱。
    Returns:
        GeoDataFrame: 更新後的公車站點 GeoDataFrame，其中 geometry 已投影到路線。
    """
    snapped_points = []

    for _, stop in stops_gdf.iterrows():
        # 找到與站點路線名稱和方向相符的路線
        matching_route = routes_gdf[(routes_gdf[route_routename_col] == stop[seq_routename_col]) & 
                                    (routes_gdf[route_direction_col] == stop[seq_direction_col])]

        if not matching_route.empty:
            # 取出該路線的 geometry
            line = matching_route.iloc[0].geometry
            # 計算站點投影到該路線的最近點
            snapped_point = line.interpolate(line.project(stop.geometry))
            snapped_points.append(snapped_point)
        else:
            # 如果沒有匹配的路線，保持原點
            snapped_points.append(stop.geometry)

    # 更新站點的 geometry
    stops_gdf = stops_gdf.copy()
    stops_gdf['geometry'] = snapped_points
    stops_gdf[seq_lat_col] = stops_gdf.geometry.y
    stops_gdf[seq_lng_col] = stops_gdf.geometry.x
    return stops_gdf

def split_routes(busroute_select, seq_select,
                 route_routename_col='RouteName',
                 route_direction_col='Direction',
                 seq_routename_col='RouteName',
                 seq_direction_col='Direction',
                 seq_seq_col='Seq',
                 seq_lat_col='Lat',
                 seq_lng_col='Lon'):
    """
    將公車路線 (busroute_select) 依據提供的站序 (seq_select) 上，分為數段的shp。
    Parameters:
        busroute_select (GeoDataFrame): 包含公車路線名稱的 GeoDataFrame。
        seq_select (DataFrame): 包含公車路線站序的 DataFrame。
        seq_routename_col (str): 路線名稱欄位名稱。
        seq_direction_col (str): 路線方向欄位名稱。
        seq_seq_col (str): 站點方向欄位名稱。
        seq_lat_col (str): 站點緯度欄位名稱。
        seq_lng_col (str): 站點經度欄位名稱。
    Returns:
        GeoDataFrame: 更新後的公車站點 GeoDataFrame，其中 geometry 已投影到路線。
    """

    output = []

    for _, route in busroute_select.iterrows():
        route_name = route[route_routename_col]
        direction = route[route_direction_col]
        geometry = route['geometry']

        # 過濾對應路線與方向的站點
        stops = seq_select[(seq_select[seq_routename_col] == route_name) & 
                           (seq_select[seq_direction_col] == direction)].sort_values(seq_seq_col)

        # 確保站點順序對應於路線
        stop_coords = [(row[seq_lng_col], row[seq_lat_col]) for _, row in stops.iterrows()]

        for i in range(len(stop_coords) - 1):
            start_point = Point(stop_coords[i])
            end_point = Point(stop_coords[i + 1])

            # 找到站點在路線中的比例位置
            start_distance = geometry.project(start_point)
            end_distance = geometry.project(end_point)

            # 提取路線幾何分段
            segment = substring(geometry, start_distance, end_distance)

            output.append({
                'RouteName': route_name,
                'Direction': direction,
                'StartSeq': stops.iloc[i][seq_seq_col],
                'EndSeq': stops.iloc[i + 1][seq_seq_col],
                'geometry': segment
            })

    return gpd.GeoDataFrame(output)
