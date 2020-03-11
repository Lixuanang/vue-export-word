# vue 纯前端导出 word

> 具体实现见`export-word.vue`文件

# 依赖

## docxtemplater

[docxtemplater](https://docxtemplater.readthedocs.io/en/latest/)使用 JSON 数据格式作为输入，可以处理 docx 和 ppt 模板。不像一些其它的工具，比如 docx.js, docx4j, python-docx 等，需要自己编写代码来生成文件，docxtemplater 只需要用户通过标签的形式编写模板，就可以生成文件。

```html
<html>
  <body>
    <button onclick="generate()">Generate document</button>
  </body>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/docxtemplater/3.17.1/docxtemplater.js"></script>
  <script src="https://unpkg.com/pizzip@3.0.6/dist/pizzip.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/1.3.8/FileSaver.js"></script>
  <script src="https://unpkg.com/pizzip@3.0.6/dist/pizzip-utils.js"></script>
  <script>
    function loadFile(url, callback) {
      PizZipUtils.getBinaryContent(url, callback)
    }
    function generate() {
      loadFile('https://docxtemplater.com/tag-example.docx', function(error, content) {
        if (error) {
          throw error
        }

        // The error object contains additional information when logged with JSON.stringify (it contains a properties object containing all suberrors).
        function replaceErrors(key, value) {
          if (value instanceof Error) {
            return Object.getOwnPropertyNames(value).reduce(function(error, key) {
              error[key] = value[key]
              return error
            }, {})
          }
          return value
        }

        function errorHandler(error) {
          console.log(JSON.stringify({ error: error }, replaceErrors))

          if (error.properties && error.properties.errors instanceof Array) {
            const errorMessages = error.properties.errors
              .map(function(error) {
                return error.properties.explanation
              })
              .join('\n')
            console.log('errorMessages', errorMessages)
            // errorMessages is a humanly readable message looking like this :
            // 'The tag beginning with "foobar" is unopened'
          }
          throw error
        }

        var zip = new PizZip(content)
        var doc
        try {
          doc = new window.docxtemplater(zip)
        } catch (error) {
          // Catch compilation errors (errors caused by the compilation of the template : misplaced tags)
          errorHandler(error)
        }

        doc.setData({
          first_name: 'John',
          last_name: 'Doe',
          phone: '0652455478',
          description: 'New Website'
        })
        try {
          // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
          doc.render()
        } catch (error) {
          // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
          errorHandler(error)
        }

        var out = doc.getZip().generate({
          type: 'blob',
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        }) //Output the document using Data-URI
        saveAs(out, 'output.docx')
      })
    }
  </script>
</html>
```

## jszip-utils

[jszip-utils](https://stuk.github.io/jszip-utils/) 提供一个 getBinaryContent(path, data)接口，path 即是文件的路径，支持 AJAX get 请求，data 为读取的文件内容。下面是官网给的例子：

```javascript
JSZipUtils.getBinaryContent('path/to/picture.png', function(err, data) {
  if (err) {
    throw err // or handle the error
  }
  var zip = new JSZip()
  zip.file('picture.png', data, { binary: true })
})
```

## file-saver

[file-saver](https://www.npmjs.com/package/file-saver)是一款适合在客户端生成文件的工具，它提供的接口 saveAs(blob, "1.docx")将会使用到。

# 安装

```
-- 安装 docxtemplater
cnpm install docxtemplater pizzip  --save

-- 安装 jszip-utils
cnpm install jszip-utils --save

-- 安装 jszip
cnpm install jszip --save

-- 安装 FileSaver
cnpm install file-saver --save
```

# 引入

```
    import docxtemplater from 'docxtemplater'
    import PizZip from 'pizzip'
    import JSZipUtils from 'jszip-utils'
    import {saveAs} from 'file-saver'

```

# 创建模板文件

> 模板文件放在项目的`public`下

我们要先创建一个模板文件，事先定义好格式和内容。docxtemplater 之前介绍到是通过标签的形式来填充数据的，简单的数据我们可以使用`{} 变量名`来实现简单的文本替换。

1. 简单的文本替换
  > 适用于单个变量

(word)模板中定义：

```
姓名：{name}
```

(vue)设置数据中定义：

```
{
    name:'lxa'
}
```

最终生成文件中显示：

```
姓名：lxa
```

2. 循环输出

  > 适用于对象、数组

(word)模板中定义：

![](https://cdn.nlark.com/yuque/0/2020/jpeg/378417/1583833454630-assets/web-upload/abbad8bf-5da6-43ad-8aac-d4ef664b8b3f.jpeg)

(vue)中定义：
```javascript
   // 对象
   price: [
        {
          receivable: "11760.00",
          priceA: "12000.00",
          priceB: "240.00",
          priceC: "0.00"
        }
      ],
      detail: [
        {
          loadPlace: "江苏省-南京市-栖霞区仙林大道111号",
          unloadPlace: "上海-上海市-青浦区徐泾东松泽大道",
          number: "T20200228194223190100067DC9",
          goodsName: "甲苯",
          price: "240",
          tons: "60吨",
          loadTime: "2020-03-10",
          loss: "200"
        }
      ],
      // 数组
      tableData: [
        {
          number: "T20200228194223190100067DC9",
          carNum: "苏A66666",
          guaNum: "挂苏A66666",
          loadTons: "60吨",
          unloadTons: "58吨",
          price: "20",
          fare: "60",
          lossTons: "300",
          compensation: "20",
          lossPrice: "40"
        }
      ]
    };
```
还有一些其它的复杂标签，比输支持条件判断，支持段落等等。详情参考[官网文档](https://docxtemplater.readthedocs.io/en/latest/)

# 实现
> 具体实现见`export-word.vue`文件
