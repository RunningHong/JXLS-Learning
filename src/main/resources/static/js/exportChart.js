/**
 * 此js是用于导出echarts统计图时使用的
 * RunningHong at 2018-12-11 13:15
 */



/**
 * 导出Echarts统计图表
 * @param mychart EChart图表的实例
 */
function exportChart(mychart) {
    // 向后台发起请求保存图片到指定目录.
    $.ajax({
        type: 'POST',
        dataType: "json",
        url: "/reportController/saveChart",
        data: {"picInfo": myChart.getDataURL()}
    })
}

/**
 * 导出页面上所有的图表
 * @param chartsDataArray 图表.getDataURL()的数组
 */
function exportAllCharts() {
    // 需要导出的图表
    var chartsJson = JSON.stringify(
            {
                "chartsJsonArray": [myChart.getDataURL(), testChart.getDataURL()]
            }
        );

    // 使用表单提交，解决ajax无法导出文件
    var form = $("<form>");
    form.attr('style', 'display:none');
    form.attr('target', '');
    form.attr('method', 'post');
    form.attr('action', '/reportController/exportAllCharts');

    var input1 = $('<input>');
    input1.attr('type', 'hidden');
    input1.attr('name', 'chartsJson');
    input1.attr('value', chartsJson);

    $('body').append(form);
    form.append(input1);

    form.submit();
    form.remove();

    // 向后台发起请求保存图片到指定目录.
    // $.ajax({
    //     type: 'POST',
    //     dataType: "json",
    //     contentType: "application/json",
    //     url: "/reportController/exportAllCharts",
    //     data: chartsDataArray
    // })
}