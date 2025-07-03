// ==UserScript==
// @name         方正教务系统成绩分项下载
// @namespace    ikaikail@ikaikail.com
// @version      1.8
// @description  期末成绩不理想？担心被穿小鞋？不用怕！这款脚本让你期末成绩和平时成绩一目了然！支持VPN环境！
// @author       iKaiKail
// @match        *://*/jwglxt/cjcx/*
// @include      *:/*/cjcx/*
// @icon         https://www.zfsoft.com/img/zf.ico
// @grant        none
// @require      https://code.jquery.com/jquery-3.6.0.min.js
// ==/UserScript==
 
(function() {
    'use strict';
 
    // 创建下载按钮
    const createDownloadButton = () => {
        return $('<button>', {
            type: 'button',
            class: 'btn btn-primary btn-sm mx-2',
            text: '导出所有成绩'
        });
    };
    // 文件下载器
    const downloadFile = (blob, filename = `成绩单_${Date.now()}.xlsx`) => {
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    };
 
    // 动态获取WebVPN基础路径
    const getBasePath = () => {
        const currentPath = window.location.pathname;
 
        // 提取路径到cjcx目录之前的部分
        const cjcxIndex = currentPath.indexOf('/cjcx/');
        if (cjcxIndex !== -1) {
            return currentPath.substring(0, cjcxIndex);
        }
 
        // 如果找不到cjcx目录，尝试提取到最后一个斜杠之前的部分
        const lastSlashIndex = currentPath.lastIndexOf('/');
        if (lastSlashIndex !== -1) {
            return currentPath.substring(0, lastSlashIndex);
        }
 
        return '/jwglxt'; // 默认路径
    };
 
    // 处理导出请求
    const handleExport = async () => {
        try {
            const xnm = document.getElementById('xnm').value;
            const xqm = document.getElementById('xqm').value;
 
            if (!xnm || !xqm) {
                throw new Error('请先选择学年和学期');
            }
 
            // 构建符合服务器要求的参数列表
            const params = new URLSearchParams([
                ['gnmkdmKey', 'N305005'],
                ['xnm', xnm],
                ['xqm', xqm],
                ['dcclbh', 'JW_N305005_GLY'],
                ...[
                    'kcmc@课程名称',
                    'xnmmc@学年',
                    'xqmmc@学期',
                    'kkbmmc@开课学院',
                    'kch@课程代码',
                    'jxbmc@教学班',
                    'xf@学分',
                    'xmcj@成绩',
                    'xmblmc@成绩分项'
                ].map(col => ['exportModel.selectCol', col]),
                ['exportModel.exportWjgs', 'xls'],
                ['fileName', '成绩单']
            ]);
 
            // 获取基础路径并构建完整URL
            const basePath = getBasePath();
            const targetUrl = `${basePath}/cjcx/cjcx_dcXsKccjList.html`;
 
            console.log('正在请求URL:', targetUrl);
 
            const response = await fetch(targetUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded'
                },
                body: params
            });
 
            if (!response.ok) {
                throw new Error(`服务器返回异常状态码: ${response.status}`);
            }
 
            const blob = await response.blob();
            downloadFile(blob);
 
            // 显示成功提示
            const btn = document.getElementById('exportAllScoresBtn');
            if (btn) {
                const originalText = btn.innerHTML;
                btn.innerHTML = ' 导出成功';
                btn.style.backgroundColor = '#218838';
                setTimeout(() => {
                    btn.innerHTML = originalText;
                    btn.style.backgroundColor = '';
                }, 2000);
            }
        } catch (error) {
            console.error('导出操作失败:', error);
 
            // 显示错误提示
            const btn = document.getElementById('exportAllScoresBtn');
            if (btn) {
                const originalText = btn.innerHTML;
                btn.innerHTML = ' 导出失败';
                btn.style.backgroundColor = '#dc3545';
                setTimeout(() => {
                    btn.innerHTML = originalText;
                    btn.style.backgroundColor = '';
                }, 2000);
            }
 
            alert(`导出失败: ${error.message}`);
        }
    };
 
    // 初始化功能
    const init = () => {
        // 移除旧按钮避免重复
        $('#exportAllScoresBtn').remove();
 
        // 创建新按钮
        const $downloadButton = createDownloadButton().click(handleExport);
 
        // 尝试多种插入位置
        const insertButton = () => {
            // 尝试插入到查询按钮旁
            const $searchButton = $('#search_go');
            if ($searchButton.length) {
                $searchButton.after($downloadButton);
                return true;
            }
 
            // 尝试插入到查询面板
            const $queryPanel = $('.panel-info:has(.panel-body)');
            if ($queryPanel.length) {
                $queryPanel.find('.panel-body').append($downloadButton);
                return true;
            }
 
            // 尝试插入到表单区域
            const $form = $('form:has(#xnm)');
            if ($form.length) {
                $form.append($downloadButton);
                return true;
            }
 
            // 最终插入到页面顶部
            $('body').prepend($downloadButton);
            return true;
        };
 
        insertButton();
    };
 
    // 增强版DOM检测
    const checkDOM = () => {
        // 确保关键元素已加载
        if ($('#xnm, #xqm').length >= 2) {
            init();
        } else {
            setTimeout(checkDOM, 500);
        }
    };
 
    // 启动检测
    setTimeout(checkDOM, 1000);
 
    // 监听AJAX请求完成
    $(document).ajaxStop(function() {
        if (!$('#exportAllScoresBtn').length) {
            setTimeout(init, 500);
        }
    });
})();
