#!/usr/bin/env node
"use strict";
const xlsx = require("node-xlsx");
const fs = require("fs");
function toJson(excelName, parent) {
  const list = xlsx.parse(excelName); // 需要 转换的excel文件
  // 数据处理 方便粘贴复制
  const data = list[0].data; // 1.读取json数据到变量暂存
  const d = {
    图层: "label",
    内容: "label",
    颜色: "colorCode",
  };
  const len = data.length;
  const columns = data[0];
  let flagIndex = null;
  let name = null;
  let lock = false;
  for (let i = 1; i < len; i++) {
    const item = data[i];
    let result = {};
    if (item[0]) {
      lock = false;
      name = item[0].replace(/^[0-9]、/, "");
    } else {
      const index = parent.children.reduce((acc, cur, index) => {
        return cur.label === name ? index : acc;
      }, -1);
      result = parent.children[index];
    }
    item.forEach((v, index, arr) => {
      const key = d[columns[index]];
      if (index === 0 && v) {
        result[key] = v.replace(/^[0-9]、/, "");
        result.icon = "() => null";
        parent.children = parent.children || [];
        parent.children.push(result);
        result.value = `${parent.value}-${parent.children.length}`;
      }

      if (index === 1) {
        result.children = result.children || [];
        const value = {
          [key]: v,
          icon: "() => null",
        };
        result.children.push(value);
        flagIndex = result.children.length - 1;
        value.value = `${parent.value}-${result.children.length}`;
        if (!arr[0] && !lock) {
          lock = true;
          for (let j = i + 1; j < len; j++) {
            if (data[j][0]) {
              break;
            }
            const vg = {
              [key]: data[j][1],
              icon: "() => null",
              [d["颜色"]]: `rgba(${data[j][2]})`,
            };
            result.children.push(vg);
            const l = result.children.length
            vg.value = `${result.children[l - 2].value.replace(/-[0-9]{1,}$/, "")}-${l}`
          }
        }
      }
      if (index === 2) {
        result.children[flagIndex][key] = `rgba(${v})`;
      }
    });
  }

  writeFile(
    "json/" + excelName.split(".")[0] + ".json",
    JSON.stringify(parent)
  ); // 输出的json文件  3.数据写入本地json文件
  function writeFile(fileName, data) {
    fs.writeFile(fileName, data, "utf-8", complete); // 文件编码格式  utf-8
    function complete(err) {
      if (!err) {
        console.log("文件生成成功"); // 终端打印这个 表示输出完成
      }
    }
  }
}
const data = [
  {
    excelName: "gjgyld.xlsx",
    parent: {
      label: "国家级公益林（地）",
      value: "12",
      icon: "() => null",
    },
  },
  {
    excelName: "ldtr.xlsx",
    parent: {
      label: "立地土壤",
      value: "13",
      icon: "() => null",
    },
  },
  {
    excelName: "fhl.xlsx",
    parent: {
      label: "防护林",
      value: "14",
      icon: "() => null",
    },
  },
];
data.forEach(({ excelName, parent }) => {
  toJson(excelName, parent);
});
