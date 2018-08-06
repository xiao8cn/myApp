var async = require ( 'async' );
var officegen = require('officegen');

var fs = require('fs');
var path = require('path');

// var themeXml = fs.readFileSync ( path.resolve ( __dirname, 'themes/testTheme.xml' ), 'utf8' );

module.exports.shengcDocx = function (options) {

	var defaultOption = {
		contractNo: '123456',
		comapayName: '测试公司'
	}

	var allOptions = Object.assign({}, defaultOption, options)

	var docx = officegen ( {
		type: 'docx',
		orientation: 'portrait',
		pageMargins: { top: 10.500, left: 10.500, bottom: 10.500, right: 10.500 }
		// The theme support is NOT working yet...
		// themeXml: themeXml
	} );
	
	// Remove this comment in case of debugging Officegen:
	// officegen.setVerboseMode ( true );
	
	docx.on ( 'error', function ( err ) {
		console.log ( err );
	});

	var table = [
		['合同编号',`${allOptions.contractNo}`],
		['签订地点',''],
		['签订日期',''],
	]
	
	var tableStyle = {
		tableColWidth: 1200,
		tableSize: 24,
		tableColor: "ada",
		tableAlign: "right",
		// align: "center",
	}
	
	console.log(docx)

	// 表格
	docx.createTable (table, tableStyle);

	// 空行
	docx.createP ({align: 'center'})
	docx.createP ({align: 'center'})
	docx.createP ({align: 'center'})
	docx.createP ({align: 'center'})
	docx.createP ({align: 'center'})

	var pObj = docx.createP({align: 'center'})
	pObj.addText('物资采购合同', {font_face: '宋体', font_size: 40, bold: true})

	// 9 空行
	docx.createP ({ align: 'center' })
	docx.createP ({ align: 'center' })
	docx.createP ({ align: 'center' })
	docx.createP ({ align: 'center' })
	docx.createP ({ align: 'center' })
	docx.createP ({ align: 'center' })
	docx.createP ({ align: 'center' })
	docx.createP ({ align: 'center' })
	docx.createP ({ align: 'center' })
	docx.createP ({ align: 'center' })
	docx.createP ({ align: 'center' })

	var pObj = docx.createP()
	pObj.addText(`      甲方（购货方）：${allOptions.comapayName}`, {font_face: '宋体', font_size: 15})
	docx.createP ({ align: 'center' })
	var pObj = docx.createP()
	pObj.addText(`      已方（销售方）：${allOptions.comapayName}`, {font_face: '宋体', font_size: 15})
	
	// 空行
	docx.createP ({ align: 'center' })
	docx.createP ({ align: 'center' })

	var pObj = docx.createP({align: "center"})
	pObj.addText(`物资采购合同`, {font_face: '宋体', font_size: 22, bold: true})

	var pObj = docx.createP()
	pObj.addText('购货方（简称甲方）：', {font_face: '宋体', font_size: 10.5, bold: true})
	pObj.addText(`${allOptions.comapayName}`, {font_face: '宋体', font_size: 10.5})
	var pObj = docx.createP()
	pObj.addText('销售方（简称乙方）：', {font_face: '宋体', font_size: 10.5, bold: true})
	pObj.addText(`${allOptions.comapayName}`, {font_face: '宋体', font_size: 10.5})
	var pObj = docx.createP()
	pObj.addText('销售方（简称乙方）：', {font_face: '宋体', font_size: 10.5})

	// 空行
	docx.createP ({ align: 'center' })

	var pObj = docx.createP()
	pObj.addText('一、物资清单', {font_face: '宋体', font_size: 10.5})

	var wuziTable = [
		['序号', '产品名称', '规格型号', '金额(元)', '计划来源']
	]

	var wuziTableStyle = {
		tableColWidth: 1200,
		tableSize: 24,
		tableColor: "ada",
		tableAlign: "center",
		borders: true
	}

	// 表格
	docx.createTable (wuziTable, wuziTableStyle)

	// 空行
	docx.createP ({ align: 'center' })

	var pObj = docx.createP()
	pObj.addText('合同金额：人民币含税金额')
	pObj.addText('       元', { bold: true, underline: true })
	pObj.addText('（含')
	pObj.addText('   % 增值税', { bold: true, underline: true })
	pObj.addText('）； 大写：')
	pObj.addText(' 人民币      圆整。', { bold: true, underline: true })

	var out = fs.createWriteStream ( 'public/tmp/out.docx' );
	
	out.on ( 'error', function ( err ) {
		console.log ( err );
	});
	
	async.parallel ([
		function ( done ) {
			out.on ( 'close', function () {
				console.log ( 'Finish to create a DOCX file.' );
				done ( null );
			});
			docx.generate ( out );
		}
	
	], function ( err ) {
		if ( err ) {
			console.log ( 'error: ' + err );
		} // Endif.
	});
}
