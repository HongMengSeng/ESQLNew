using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace ESQLNew.Core
{
    public static class ColumnMapStore
    {
        private static bool _lastLoadCorrupt;

        public static bool LastLoadCorrupt
        {
            get
            {
                return _lastLoadCorrupt;
            }
        }
        public static string ConfigPath
        {
            get
            {
                return AppPaths.ColumnMapsFile;
            }
        }

        public static Dictionary<string, Dictionary<string, string>> BuiltinMap
        {
            get
            {
                return new Dictionary<string, Dictionary<string, string>>
                {
                    { "sx_qms_repair_info_history", new Dictionary<string, string>
                        {
                            { "序号", "seq_no" },
                            { "担当", "person_in_charge" },
                            { "维修单号", "repair_order_no" },
                            { "次数", "order_times" },
                            { "重复单", "is_duplicate" },
                            { "地区", "region" },
                            { "省份", "province" },
                            { "城市", "city" },
                            { "大小型机", "machine_size" },
                            { "店铺编号", "shop_code" },
                            { "店铺名称", "shop_name" },
                            { "分类", "category" },
                            { "系列", "series" },
                            { "室内机型", "indoor_model" },
                            { "室外机型", "outdoor_model" },
                            { "机型", "model" },
                            { "机型年度", "model_year" },
                            { "机型年份", "year_label" },
                            { "室内机编号", "indoor_serial_no" },
                            { "室外机编号", "outdoor_serial_no" },
                            { "购买日期", "purchase_date" },
                            { "安装日期", "installation_date" },
                            { "故障日期", "failure_date" },
                            { "使用期限（天）", "service_life_days" },
                            { "年限", "age_limit" },
                            { "销售店", "sales_shop" },
                            { "安装店编码", "installation_shop_code" },
                            { "安装店", "installation_shop" },
                            { "受理时间", "acceptance_time" },
                            { "指定上门时间", "scheduled_visit_time" },
                            { "实际上门时间", "actual_visit_time" },
                            { "完成时间", "completion_time" },
                            { "首次维修日期", "first_repair_date" },
                            { "上次维修日期", "last_repair_date" },
                            { "维修次数", "repair_times" },
                            { "回访服务态度", "follow_up_attitude" },
                            { "回访满意度", "follow_up_satisfaction" },
                            { "其它意见", "other_opinion" },
                            { "担当审核", "person_in_charge_review" },
                            { "担当备注", "person_in_charge_remark" },
                            { "担当审核日期", "person_in_charge_review_date" },
                            { "回访", "follow_up" },
                            { "回访员", "follow_up_staff" },
                            { "回访日期", "follow_up_date" },
                            { "回访备注", "follow_up_remark" },
                            { "部长审核", "department_head_review" },
                            { "部长备注", "department_head_remark" },
                            { "部长审核日期", "department_head_review_date" },
                            { "单据状态", "document_status" },
                            { "交通距离", "travel_distance" },
                            { "上传月", "upload_month" },
                            { "上传时间", "upload_time" },
                            { "店铺备注", "shop_remark" },
                            { "故障现象", "failure_phenomenon" },
                            { "故障症状", "failure_symptom" },
                            { "故障原因", "failure_cause" },
                            { "处理方法", "handling_method" },
                            { "部件编码", "component_code" },
                            { "部件名称", "component_name" },
                            { "部件价格", "component_price" },
                            { "是否返厂机", "is_return_factory_machine" },
                            { "受理单号", "acceptance_no" },
                            { "是否换件", "is_replaced" },
                            { "转录保修单", "transfer_warranty" },
                            { "结算状态", "settlement_status" },
                            { "结算期数", "settlement_period" },
                            { "基本维修费", "basic_repair_fee" },
                            { "交通费", "transportation_fee" },
                            { "速度费", "speed_fee" },
                            { "信息费", "information_fee" },
                            { "上门费", "visit_fee" },
                            { "完成费", "completion_fee" },
                            { "扣罚款", "penalty_amount" },
                            { "扣罚款原因", "penalty_reason" },
                            { "调整费用", "adjustment_fee" },
                            { "调整备注", "adjustment_remark" },
                            { "结算金额", "settlement_amount" },
                            { "合计", "total_amount" },
                            { "项目", "item" },
                            { "室内外区分", "indoor_outdoor_type" },
                            { "期限", "warranty_period" }
                        } }
                };
            }
        }

        public static Dictionary<string, Dictionary<string, string>> Load(string path)
        {
            _lastLoadCorrupt = false;
            try
            {
                if (!File.Exists(path))
                {
                    var init = BuiltinMap;
                    Save(path, init);
                    return init;
                }
                var text = File.ReadAllText(path);
                var parsed = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, string>>>(text);
                if (parsed != null) return parsed;
                return BuiltinMap;
            }
            catch
            {
                _lastLoadCorrupt = true;
                return BuiltinMap;
            }
        }

        public static void Save(string path, Dictionary<string, Dictionary<string, string>> maps)
        {
            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            File.WriteAllText(path, JsonConvert.SerializeObject(maps, Formatting.Indented));
        }

        public static Dictionary<string, string> GetMap(string path, string table)
        {
            var maps = Load(path);
            Dictionary<string, string> result;
            if (maps.TryGetValue(table, out result))
                return result;
            return new Dictionary<string, string>();
        }

        public static Dictionary<string, string> BuildMap(IList<string> excelHeaders, IList<string> selectedFields)
        {
            var result = new Dictionary<string, string>();
            if (excelHeaders == null || selectedFields == null || excelHeaders.Count != selectedFields.Count)
                return result;
            for (int i = 0; i < excelHeaders.Count; i++)
            {
                var header = excelHeaders[i] == null ? "" : excelHeaders[i].Trim();
                var field = selectedFields[i] == null ? "" : selectedFields[i].Trim();
                if (header.Length == 0 || field.Length == 0) continue;
                result[header] = field;
            }
            return result;
        }
    }
}