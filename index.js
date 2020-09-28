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
  let flagValue = parent.value;
  for (let i = 1; i < len; i++) {
    const item = data[i];
    const result = {};
    item.forEach((v, index) => {
      const key = d[columns[index]];
      if (index === 0) {
        result[key] = v.replace(/^[0-9]、/, "");
        result.icon = "() => null";
        parent.children = parent.children || [];
        parent.children.push(result);
        result.value = `${parent.value}-${parent.children.length}`;
        flagValue = `${parent.value}-${parent.children.length}`;
      }
      if (index === 1) {
        result.children = result.children || [];
        const value = {
          [key]: v,
          icon: "() => null",
        };
        result.children.push(value);
        flagIndex = result.children.length - 1;
        value.value = `${flagValue}-${result.children.length}`;
        flagValue = null;
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
