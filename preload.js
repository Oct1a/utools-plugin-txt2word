var fs = require('fs');
var $path = require('path');
const { TextRun, Packer, Document, Paragraph } = require("docx");

function isDir(dir) {
  fs.exists(dir, function(exists) {
    if (!exists) {
      fs.mkdir(dir, err => {})
    }
  });
}

window.exports = {
  "convert": {
    mode: "none",
    args: {
      enter: (action, callbackSetList) => {
        utools.hideMainWindow()
        let files = action.payload
          // 遍历选中文件
        for (const element of files) {
          if (element.isFile) {
            let { name, path } = element
            let folder_path = $path.dirname(path)
            let end_path = $path.join(folder_path, "转为word")
            let out_name = name.replace("txt", "docx")
            isDir(end_path)
              // 读取文件
            fs.readFile(path, 'utf-8', function(err, data) {
              if (err) {
                console.log(err);
              } else {
                let process = data.replace(/(\r\n)|(\n)/g, '\t');
                let res = []
                process.split("\t").forEach(v => {
                  res.push(new Paragraph({
                    children: [
                      new TextRun(v),
                    ],
                  }))
                })
                const doc = new Document({
                  sections: [{
                    properties: {},
                    children: res
                  }],
                });
                // 写入word文件
                Packer.toBuffer(doc).then((buffer) => {
                  fs.writeFileSync($path.join(end_path, out_name), buffer)
                });
              }
            })
          }
        }
        utools.showNotification('全部转换完成！')
      }
    }
  }
}