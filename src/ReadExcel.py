import openpyxl


path = "/Users/gaoxing/pythod_project/generate_insert/excel_source/"
fileName = "0.xlsx"
cityId = 68
cityAreaId = 4
sqlFile = "add1.sql"


def go():
    read_excel()


def read_excel():
    wb = openpyxl.load_workbook("%s%s" % (path, fileName))
    sheets = wb.get_sheet_names()
    sheet = wb.get_sheet_by_name(sheets[0])
    size = sheet.max_row
    prefix = ["B", "C", "D", "E", "G", "H", "I"]
    for i in range(3, size+1):
        lists = []
        for index in prefix:
            lists.append(sheet['%s%s' % (index, i)].value)
        make_insert(lists)


def make_insert(lists):
    if not lists[0]:
        return
    model_id = 1;
    if "7S" in lists[6]:
        model_id = 2;

    car_info = {"car_sn": lists[0], "tbox_sn": lists[1], "iccid": lists[2],"factory_card_no": lists[3],"car_plate": lists[4],"car_frame_sn": lists[5], "modelId": model_id}
    car_info_insert_sql = generate_car_info_sql(car_info)
    print(car_info_insert_sql)

    file = open("%s%s" % (path, sqlFile), "a")
    file.writelines("%s %s" % (car_info_insert_sql, "\n"))

    car_stock_insert_sql = generate_car_stock_sql(car_info)
    print(car_stock_insert_sql)
    file.writelines("%s %s" % (car_stock_insert_sql, "\n"))

    tbox_stock_insert_sql = generate_tbox_stock_sql(car_info)
    print(tbox_stock_insert_sql)
    file.writelines("%s %s" % (tbox_stock_insert_sql, "\n"))

    car_geo_info_insert_sql = generate_car_geo_info_sql(car_info)
    print(car_geo_info_insert_sql)
    file.writelines("%s %s" % (car_geo_info_insert_sql, "\n"))


def generate_car_info_sql(car_info):
    car_info_insert_sql = "INSERT INTO `car_info` (`car_sn`, `tbox_sn`, `iccid`, `card_no`, `longitude`, `latitude`, `left_mileage`, `total_milleage`, `soc`, `vol`, `s_vol`, `gprs`, `alarm`, `net_status`, `net_server_time`, `net_tbox_time`, `work_status`, `e_status`, `acc_status`, `d_status`, `lock_status`, `wd_status`, `light_status`, `charge_status`, `last_report_time`, `bt_mac`, `bt_token`, `bt_tbox_time`, `bt_server_time`, `create_time`, `update_time`) VALUES ('%s', '%s', '%s', '%s', 0.00, 0.00, 0, 0, 0.00, 0.00, 0.00, 0, 0, 0, '0000-00-00 00:00:00', '0000-00-00 00:00:00', 0, 0, 0, 0, 0, 0, 0, 0, '0000-00-00 00:00:00', '', '', '0000-00-00 00:00:00', '0000-00-00 00:00:00', now(), '0000-00-00 00:00:00');" % (car_info["car_sn"], car_info["tbox_sn"], car_info["iccid"], car_info["factory_card_no"])

    return car_info_insert_sql


def generate_car_stock_sql(car_info):
    car_stock_insert_sql = "INSERT INTO `car_stock` (`car_sn`, `car_frame_sn`, `status`, `car_plate`, `model_id`, `city_id`, `create_time`, `update_time`) VALUES ('%s', '%s', 2, '%s', %s, %s, now(), '0000-00-00 00:00:00');" % (car_info["car_sn"], car_info["car_frame_sn"], car_info["car_plate"], car_info["modelId"], cityId)

    return car_stock_insert_sql


def generate_tbox_stock_sql(car_info):
    car_stock_insert_sql = "INSERT INTO `tbox_stock` (`tbox_sn`, `pro_ver`, `factory_iccid`, `factory_card_no`, `factory_soft_ver`, `factory_hard_ver`, `create_time`, `update_time`) VALUES ('%s', '', '%s', '%s', '040000100', '', now(), '0000-00-00 00:00:00');" % (car_info["tbox_sn"], car_info["iccid"], car_info["factory_card_no"])

    return car_stock_insert_sql


def generate_car_geo_info_sql(car_info):
    car_stock_insert_sql = "INSERT INTO `car_geo_info` (`car_sn`, `city_id`, `last_order_start_time`, `last_launch_time`, `last_recycle_time`, `city_area_id`, `create_time`, `update_time`) VALUES ('%s', %s, '0000-00-00 00:00:00', '0000-00-00 00:00:00', '0000-00-00 00:00:00', %s, now(), '0000-00-00 00:00:00');" % (car_info["car_sn"], cityId, cityAreaId)

    return car_stock_insert_sql


go()
