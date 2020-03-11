<!--  -->

<template>
  <div class="page-container">
    <button @click="generate">Generate document</button>
  </div>
</template>

<script>
import docxtemplater from 'docxtemplater'
import PizZip from 'pizzip'
import JSZipUtils from 'jszip-utils'
import { saveAs } from 'file-saver'
export default {
  components: {},

  props: {},

  data() {
    return {
      price: [
        {
          receivable: '11760.00',
          priceA: '12000.00',
          priceB: '240.00',
          priceC: '0.00'
        }
      ],
      detail: [
        {
          loadPlace: '江苏省-南京市-栖霞区仙林大道111号',
          unloadPlace: '上海-上海市-青浦区徐泾东松泽大道',
          number: 'T20200228194223190100067DC9',
          goodsName: '甲苯',
          price: '240',
          tons: '60吨',
          loadTime: '2020-03-10',
          loss: '200'
        }
      ],
      tableData: [
        {
          number: 'T20200228194223190100067DC9',
          carNum: '苏A66666',
          guaNum: '挂苏A66666',
          loadTons: '60吨',
          unloadTons: '58吨',
          price: '20',
          fare: '60',
          lossTons: '300',
          compensation: '20',
          lossPrice: '40'
        }
      ]
    }
  },
  computed: {},
  watch: {},

  mounted() {},
  methods: {
    generate() {
      let that = this
      // 读取并获得模板文件的二进制内容
      JSZipUtils.getBinaryContent('/module.docx', function(error, content) {
        // model.docx是模板。我们在导出的时候，会根据此模板来导出对应的数据
        // 抛出异常
        if (error) {
          throw error
        }

        // 创建一个PizZip实例，内容为模板的内容
        let zip = new PizZip(content)
        // 创建并加载docxtemplater实例对象
        let doc = new docxtemplater().loadZip(zip)
        // 设置模板变量的值
        doc.setData({
          table: that.tableData,
          price: that.price,
          detail: that.detail
        })

        try {
          // 用模板变量的值替换所有模板变量
          doc.render()
        } catch (error) {
          // 抛出异常
          let e = {
            message: error.message,
            name: error.name,
            stack: error.stack,
            properties: error.properties
          }
          console.log(JSON.stringify({ error: e }))
          throw error
        }

        // 生成一个代表docxtemplater对象的zip文件（不是一个真实的文件，而是在内存中的表示）
        let out = doc.getZip().generate({
          type: 'blob',
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        })
        // 将目标文件对象保存为目标类型的文件，并命名
        saveAs(out, '对账.docx')
      })
    }
  }
}
</script>
