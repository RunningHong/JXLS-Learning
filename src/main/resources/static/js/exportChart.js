/**
 * 此js是用于导出echarts统计图时使用的
 * RunningHong at 2018-12-11 13:15
 */



/**
 * 导出Echarts统计图表
 * @param mychart:EChart图表的实例
 */
function exportChart(mychart) {
    // 向后台发起请求保存图片到指定目录.
    $.ajax({
        type: 'POST',
        dataType: "json",
        url: "/chart/saveChartImage",
        data: {"picInfo": myChart.getDataURL()},
        success: function () {
            console.log('导出Echarts统计图表成功。');
        }
    })
}