using System;
using System.Collections.Generic;
using System.Globalization;

namespace ESQLNew.Import
{
    public static class ColumnMapper
    {
        private static readonly Dictionary<string, string> FieldMap = new Dictionary<string, string>
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
        };

        public static IList<ColumnMapping> Map(IList<string> excelHeaders, IList<ColumnInfo> tableColumns)
        {
            var result = new List<ColumnMapping>(excelHeaders.Count);
            foreach (var h in excelHeaders)
            {
                ColumnInfo matched = null;
                string mappedField = null;
                string key = h.Trim();
                if (FieldMap.TryGetValue(key, out mappedField))
                {
                    foreach (var c in tableColumns)
                    {
                        if (string.Equals(c.Name.Trim('`'), mappedField, StringComparison.OrdinalIgnoreCase))
                        {
                            matched = c;
                            break;
                        }
                    }
                }
                if (matched == null)
                {
                    foreach (var c in tableColumns)
                    {
                        if (string.Equals(c.Name.Trim('`'), h.Trim(), StringComparison.OrdinalIgnoreCase))
                        {
                            matched = c;
                            break;
                        }
                    }
                }
                result.Add(new ColumnMapping
                {
                    ExcelColumn = h,
                    TableField = matched != null ? matched.Name : null,
                    Matched = matched != null,
                    TableColumn = matched
                });
            }
            return result;
        }

        public static object ConvertValue(object raw, ColumnInfo col)
        {
            if (raw == null) return null;
            string s = raw as string;
            if (s != null && s.Trim().Length == 0) return null;

            string type = (col.DataType ?? "").ToLowerInvariant();

            if (type == "datetime" || type == "timestamp" || type == "date" || type == "time")
            {
                var dt = raw as DateTime?;
                if (dt.HasValue) return dt.Value;
                var oa = raw as double?;
                if (!oa.HasValue) oa = raw as float?;
                if (oa.HasValue)
                {
                    try
                    {
                        return DateTime.FromOADate(oa.Value);
                    }
                    catch (ArgumentException)
                    {
                        throw new FormatException("无法解析日期: " + raw);
                    }
                }
                if (s != null)
                {
                    DateTime parsed;
                    if (DateTime.TryParse(s.Trim(), out parsed))
                        return parsed;
                }
                throw new FormatException("无法解析日期: " + s);
            }

            if (type == "int" || type == "bigint" || type == "smallint" || type == "tinyint" || type == "mediumint")
            {
                var n = raw as int?;
                if (n.HasValue) return (long)n.Value;
                var l = raw as long?;
                if (l.HasValue) return l.Value;
                var d = raw as double?;
                if (d.HasValue) return (long)d.Value;
                if (s != null)
                {
                    long parsed;
                    if (long.TryParse(s.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out parsed))
                        return parsed;
                    double dd;
                    if (double.TryParse(s.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out dd))
                        return (long)dd;
                }
                throw new FormatException("无法解析整数: " + s);
            }

            if (type == "decimal" || type == "double" || type == "float")
            {
                var d = raw as double?;
                if (d.HasValue) return Convert.ToDecimal(d.Value);
                var dec = raw as decimal?;
                if (dec.HasValue) return dec.Value;
                if (s != null)
                {
                    decimal parsed;
                    if (decimal.TryParse(s.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out parsed))
                        return parsed;
                }
                throw new FormatException("无法解析数字: " + s);
            }

            if (col.MaxLength > 0 && s != null && s.Length > col.MaxLength)
                throw new FormatException("超出字段长度 " + col.MaxLength + ": " + s);

            return s ?? raw.ToString();
        }
    }
}