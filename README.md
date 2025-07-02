# 方正教务系统成绩分项下载
这个脚本可以在使用方正教务系统的学校网站中添加一个"导出所有成绩"功能，特别针对成绩查询页面。下面首先是功能展示，其次我将详细解释脚本的原理和工作流程：

### 功能展示
在方正教务系统中找到成绩查询界面
将学年，学期全部改为你要查询的学期，点击查询，此时不会显示成绩
然后直接点击右边“导出所有成绩”按钮
点击后，浏览器会自动下载一个xls文件，里面就是所有的成绩
如需选择导出其他学期成绩重新查询导出即可
如果用不了请邮箱反馈或者添加微信wxuid_ikaikail联系

<img width="100" alt="1" src="https://github.com/user-attachments/assets/11c73548-d232-4b3b-bbca-6e1ac5439c4e" />
<img width="100" alt="2" src="https://github.com/user-attachments/assets/5ccdd3de-1020-408a-b2fb-2d8a44407656" />
<img width="100" alt="3" src="https://github.com/user-attachments/assets/4f3274e2-3063-47c6-8921-2494b7e40ce7" />


### 脚本核心功能
脚本的核心功能是添加一个按钮，允许用户一键导出包含详细成绩分项信息的Excel文件（.xls格式），而不是页面上展示的简化成绩信息。


### 脚本原理详解

#### 1. 元信息声明（Metadata）
```javascript
// ==UserScript==
// @name         方正教务系统成绩分项下载
// @namespace    ikaikail@ikaikail.com
// @version      1.3
// ...
// @match        *://jwgl.sxzy.edu.cn/jwglxt/cjcx/*
// @match        *://jwgl.fafu.edu.cn/jwglxt/cjcx/*
// ...（多个学校匹配规则）...
// @match        *://jw.gzist.edu.cn/jwglxt/xtgl/cjcx/*
// @require      https://code.jquery.com/jquery-3.6.0.min.js
// ==/UserScript==
```
- **@match**：定义了脚本生效的URL模式，匹配多个使用方正教务系统的学校
- **@require**：加载jQuery库，用于DOM操作
- **@version**：脚本版本号
- **@downloadURL/@updateURL**：提供脚本的自动更新功能

#### 2. 创建下载按钮
```javascript
const createDownloadButton = () => {
    return $('<button>', {
        type: 'button',
        class: 'btn btn-primary btn-sm mx-2',
        text: '导出所有成绩'
    });
};
```
- 使用jQuery创建一个按钮元素
- 设置按钮样式为`btn btn-primary`（Bootstrap样式）
- 按钮文本为"导出所有成绩"

#### 3. 文件下载功能
```javascript
const downloadFile = (blob, filename = `成绩单_${Date.now()}.xlsx`) => {
    // 创建临时下载链接
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    
    // 触发下载
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    // 清理资源
    URL.revokeObjectURL(url);
};
```
- 接收服务器返回的文件数据（Blob对象）
- 创建临时下载链接并触发下载
- 文件名包含时间戳确保唯一性
- 下载完成后清理临时资源

#### 4. 导出请求处理（核心）
```javascript
const handleExport = async () => {
    try {
        // 获取学年和学期值
        const xnm = document.getElementById('xnm').value;
        const xqm = document.getElementById('xqm').value;

        // 构建请求参数
        const params = new URLSearchParams([
            ['gnmkdmKey', 'N305005'], // 功能模块标识
            ['xnm', xnm], // 学年
            ['xqm', xqm], // 学期
            ['dcclbh', 'JW_N305005_GLY'], // 导出模板标识
            ...[ // 指定导出的列
                'kcmc@课程名称',
                'xnmmc@学年',
                // ...其他列...
                'xmblmc@成绩分项' // 关键：成绩分项信息
            ].map(col => ['exportModel.selectCol', col]),
            ['exportModel.exportWjgs', 'xls'], // 导出格式为Excel
            ['fileName', '成绩单'] // 文件名
        ]);

        // 发送POST请求
        const response = await fetch('/jwglxt/cjcx/cjcx_dcXsKccjList.html', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            },
            body: params
        });

        // 处理响应
        if (!response.ok) throw new Error(`服务器返回异常状态码: ${response.status}`);
        
        const blob = await response.blob();
        downloadFile(blob);
    } catch (error) {
        // 错误处理
        console.error('导出操作失败:', error);
        alert(`导出失败: ${error.message}`);
    }
};
```
**关键点：**
1. 获取当前选择的学年(`xnm`)和学期(`xqm`)
2. 构建符合教务系统要求的POST参数：
   - `gnmkdmKey`：标识成绩查询功能模块
   - `dcclbh`：指定导出模板（包含成绩分项信息的模板）
   - `exportModel.selectCol`：明确请求导出包含"成绩分项"(xmblmc)的完整数据
3. 请求路径：`/jwglxt/cjcx/cjcx_dcXsKccjList.html`（教务系统标准导出接口）
4. 设置导出格式为Excel(`xls`)

#### 5. 初始化与DOM检测
```javascript
// 初始化功能
const init = () => {
    // 查找查询按钮
    const $searchButton = $('#search_go');
    
    // 创建并配置下载按钮
    const $downloadButton = createDownloadButton()
        .click(handleExport)
        .prop('title', '导出包含所有成绩分项的完整数据');

    // 插入按钮到页面
    $searchButton.length > 0 ? 
        $searchButton.after($downloadButton) : 
        $('body').prepend($downloadButton);
};

// 等待页面元素加载
const checkDOM = () => {
    if (document.getElementById('xnm') && document.getElementById('xqm')) {
        init(); // 初始化
    } else {
        setTimeout(checkDOM, 300); // 继续等待
    }
};

// 启动检测
checkDOM();
```
- **DOM检测**：定期检查学年/学期选择框是否加载完成
- **按钮插入**：找到查询按钮后插入导出按钮
- **用户提示**：按钮添加title属性说明功能

### 为什么需要这个脚本？

1. **解决信息不全问题**：
   - 教务系统前台只显示总分
   - 后台有详细分项成绩（平时成绩、期中、期末等）
   - 脚本通过特定参数请求完整数据

2. **绕过界面限制**：
   - 普通界面不提供分项成绩导出
   - 脚本直接调用系统隐藏的导出接口

3. **标准化操作**：
   - 不同学校使用相同的方正系统
   - 脚本适配多个学校的URL模式

### 工作流程总结

1. **加载与检测**：脚本加载后等待页面关键元素（学年/学期选择框）
2. **添加按钮**：在查询按钮旁添加"导出所有成绩"按钮
3. **用户交互**：用户点击导出按钮
4. **构建请求**：
   - 获取当前学年/学期
   - 构建包含"成绩分项"参数的请求
5. **发送请求**：向教务系统后台发送导出请求
6. **处理响应**：接收Excel文件并触发下载
7. **错误处理**：捕获并显示可能的错误

### 技术亮点

1. **渐进式检测**：使用`setTimeout`轮询确保DOM加载完成
2. **参数精准构造**：通过`URLSearchParams`构建符合服务器要求的参数
3. **跨校兼容**：通过多个`@match`规则适配不同学校
4. **用户体验**：按钮样式与原生界面保持一致
5. **错误处理**：完整的try-catch错误捕获机制

这个脚本本质上是利用了教务系统已有的导出接口，但通过特定参数请求了包含详细分项成绩的数据，这些数据在普通界面中是不可见的，但对学生了解成绩构成非常重要。
