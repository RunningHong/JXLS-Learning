/**
 * 此js是用于导出echarts统计图时使用的
 * RunningHong at 2018-12-11 13:15
 */

// 参数集，用于传参
var params = {"testParam":"ddd"};

// 导出信息数组,存放需要导出的图表
var exportChartArray = [];


/**
 * 向参数集添加参数
 * @param addParams 对象形式的
 */
function addParams(addParams) {
    for(var param in addParams) {
        params[param] = addParams[param];
    }
}

/**
 * 添加需要导出的统计图表
 * @param charts 这里传入需要导出的charts，参数不限量(举例：myChart)
 * 参数是类似【var myChart = echarts.init(document.getElementById('main'));】的echarts已初始化了的值
 */
function addExportCharts() {
    exportChartArray = []; // 初始化

    // 添加图片导出信息（信息为base64压缩后的结果）
    for (var i = 0; i < arguments.length; i++) {
        exportChartArray.push(arguments[i].getDataURL());
    }
}

/**
 * 报表导出
 * @param charts 这里传入需要导出的charts，参数不限量(举例：charts1)
 * 参数需是类似【var myChart = echarts.init(document.getElementById('main'));】的echarts初始化了的值
 */
function reportExport(type) {
    addParams({"modelFileKey": "测试报表ID"});
    addParams({"chartsArray": exportChartArray}); // 导出图片信息
    // addParams({"exportType": "信息导出"}); // 导出类型【信息导出、图表导出、合并导出】
    // addParams({"exportFileType": "pdf"}); // 导出文件类型【xls, pdf】

    if (type == 1) {
        addParams({"exportType": "信息导出"});
        addParams({"exportFileType": "xls"});
    } else if (type == 2) {
        addParams({"exportType": "图表导出"});
        addParams({"exportFileType": "xls"});
    } else if (type == 3) {
        addParams({"exportType": "合并导出"});
        addParams({"exportFileType": "xls"});
    } else if (type == 4) {
        addParams({"exportType": "信息导出"});
        addParams({"exportFileType": "pdf"});
    }

    exportRequest(params);
}



/**
 * 导出请求，从此方法请求后台进行导出
 * @param exportParams 导出所需参数
 */
function exportRequest(exportParams) {
    // 请求的路径
    var url = "/reportController/reportExport";

    // 使用表单提交，解决ajax无法导出文件
    var form = $("<form>");
    form.attr('style', 'display:none');
    form.attr('target', '_blank'); // 打开新窗口进行文件导出
    form.attr('method', 'post');
    form.attr('action', url);
    var input1 = $('<input>');
    input1.attr('type', 'hidden');
    input1.attr('name', 'params'); // 后台可以使用request接受此参数
    input1.attr('value', JSON.stringify(exportParams));
    $('body').append(form);
    form.append(input1);
    form.submit();
    form.remove();

    window.opener=null; // 关闭文件导出窗口
    window.open('','_self');
    window.close();
}

