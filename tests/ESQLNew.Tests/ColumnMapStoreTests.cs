using System;
using System.Collections.Generic;
using System.IO;
using ESQLNew.Core;
using Xunit;

namespace ESQLNew.Tests
{
    public class ColumnMapStoreTests
    {
        private static string TempPath()
        {
            return Path.Combine(Path.GetTempPath(), "colmap_" + Guid.NewGuid().ToString("N") + ".json");
        }

        [Fact]
        public void BuiltinMap_RepairTableGoldenFixture()
        {
            var expected = new Dictionary<string, string>
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
            var map = ColumnMapStore.BuiltinMap;
            Assert.True(map.ContainsKey("sx_qms_repair_info_history"));
            Assert.Equal(expected, map["sx_qms_repair_info_history"]);
        }

        [Fact]
        public void SaveThenLoad_RoundTrip()
        {
            var path = TempPath();
            try
            {
                var maps = new Dictionary<string, Dictionary<string, string>>
                {
                    { "my_table", new Dictionary<string, string> { { "订单号", "order_no" } } }
                };
                ColumnMapStore.Save(path, maps);
                var loaded = ColumnMapStore.Load(path);
                Assert.True(loaded.ContainsKey("my_table"));
                Assert.Equal("order_no", loaded["my_table"]["订单号"]);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void Load_MissingFile_InitializesWithBuiltin()
        {
            var path = TempPath();
            try
            {
                var loaded = ColumnMapStore.Load(path);
                Assert.True(loaded.ContainsKey("sx_qms_repair_info_history"));
                Assert.True(File.Exists(path));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void Load_CorruptFile_FallsBackToBuiltinAndFlagsCorrupt()
        {
            var path = TempPath();
            try
            {
                File.WriteAllText(path, "{ not valid json");
                var loaded = ColumnMapStore.Load(path);
                Assert.True(loaded.ContainsKey("sx_qms_repair_info_history"));
                Assert.True(ColumnMapStore.LastLoadCorrupt);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void Load_SuccessfulFile_ResetsLastLoadCorrupt()
        {
            var path = TempPath();
            try
            {
                File.WriteAllText(path, "{ not valid json");
                ColumnMapStore.Load(path);
                Assert.True(ColumnMapStore.LastLoadCorrupt);
                var maps = new Dictionary<string, Dictionary<string, string>>
                {
                    { "my_table", new Dictionary<string, string> { { "订单号", "order_no" } } }
                };
                ColumnMapStore.Save(path, maps);
                ColumnMapStore.Load(path);
                Assert.False(ColumnMapStore.LastLoadCorrupt);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void GetMap_MissingTable_ReturnsEmpty()
        {
            var path = TempPath();
            try
            {
                ColumnMapStore.Save(path, new Dictionary<string, Dictionary<string, string>>());
                var map = ColumnMapStore.GetMap(path, "no_such_table");
                Assert.NotNull(map);
                Assert.Empty(map);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void BuildMap_CombinesHeadersAndFields_SkipsEmpty()
        {
            var headers = new List<string> { "序号", "担当", "维修单号" };
            var fields = new List<string> { "seq_no", "", null };
            var map = ColumnMapStore.BuildMap(headers, fields);
            Assert.Equal(1, map.Count);
            Assert.Equal("seq_no", map["序号"]);
        }

        [Fact]
        public void BuildMap_TrimsHeaderAndSkipsBlankHeader()
        {
            var headers = new List<string> { " 序号 ", "", "  " };
            var fields = new List<string> { "seq_no", "order_no", "x" };
            var map = ColumnMapStore.BuildMap(headers, fields);
            Assert.Equal(1, map.Count);
            Assert.Equal("seq_no", map["序号"]);
        }

        [Fact]
        public void BuildMap_LengthMismatch_ReturnsEmpty()
        {
            var map = ColumnMapStore.BuildMap(new List<string> { "a" }, new List<string> { "x", "y" });
            Assert.Empty(map);
        }
    }
}
