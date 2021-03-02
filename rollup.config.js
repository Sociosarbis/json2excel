export default {
  input: "js/worker.js",
  output:{
  	format: "iife",
  	file: "json2excel.worker.js"
  },
  moduleContext:() => "self"
};