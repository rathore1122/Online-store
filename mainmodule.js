var tempData=[],localData2=[], localData=[], mainData=[[]], staticData=[];

var fs = require('fs'); 
var parse = require('csv-parse');
var ExcelBuilder=require('excel-builder');

var csvData=[];
fs.createReadStream('farzi222.csv')
    .pipe(parse({delimiter: ';',trim:true,}))
    .on('data', function(csvrow) {
        //do something with csvrow
        csvData.push(csvrow);        
    })
    .on('end',function() {
      //do something wiht csvData
      //console.log(csvData);
	  for(var i=0;i<csvData.length;i++){
		  localData.push((csvData[i][0]).trim());
	  }
	
    });
	fs.createReadStream('farzi2222.csv')
    .pipe(parse({delimiter: ';',trim:true,}))
    .on('data', function(csvrow) {
        localData2.push(csvrow); 
    }).on('end',function(e){
		//console.log("HIi",localData2);
		pk();
	});
	
	var cellsDefination={};
		for(var i=0, j=65,k=64;i<702;i++){
			if(i<26 && j<91){
				cellsDefination[i]=String.fromCharCode(j++);
			}
			else if(j<91){
				cellsDefination[i]=String.fromCharCode(k)+String.fromCharCode(j++);
			}
			else{
				j=65;
				k++;
				i--;
			}
		}
	console.log('Cells Defination Over');
		
		
		//mainData.push(largeData);
		function pk(){
			console.log('Entered');
			
			var workbook = ExcelBuilder.Builder.createWorkbook(), stylesheet = workbook.getStyleSheet(); 	
			var sheet1 = workbook.createWorksheet({
					name: 'Sheet1'
				});
			sheet1.setRowInstructions(1, {
					height: 25
				});
			sheet1.mergeCells('A1','I3');
			
			var headings = stylesheet.createFormat({
							"font": {
								"size": 10,
								"color":"000000"
							},
							 "alignment": {
								"vertical": "left"
							}
						});
			var headings1 = stylesheet.createFormat({
							"font": {
								"size": 24,
								"color":"ffff00",
								"weight":'bolder'
							},
							
							"fill": {
								"type": 'pattern',
								"patternType": 'solid',
								"fgColor": '002850'
							},
							 "alignment": {
								"vertical": "top",
								"horizontal":"left"
							}
						});
			
			var headings2 = stylesheet.createFormat({
							
							"font": {
								"size": 18,
								"color":"000000",
								"weight":'bolder'
							},
							"fill": {
								"type": 'pattern',
								"patternType": 'solid',
								"fgColor": 'd8e4bc'
							},
							 "alignment": {
								"horizontal": "center"
							}
						});
			var headings3 = stylesheet.createFormat({
							
							"font": {
								"size": 11,
								"color":"000000",
								"weight":'bold'
							},
							"fill": {
								"type": 'pattern',
								"patternType": 'solid',
								"fgColor": 'd8e4bc'
							},
							 "alignment": {
								"horizontal": "center"
							}
						});
			var headings4 = stylesheet.createFormat({
							"font": {
								"size": 11,
								"color":"ffffff",
								"weight":'bold'
							},
							
							"fill": {
								"type": 'pattern',
								"patternType": 'solid',
								"fgColor": '002850'
							},
							 "alignment": {
								"vertical": "left"
							}
						});
			var headings5 = stylesheet.createFormat({
							
							"font": {
								"size": 10,
								"color":"000000",
								"weight":'bold'
							},
							
							"fill": {
								"type": 'pattern',
								"patternType": 'solid',
								"fgColor": 'd8e4bc'
							},
							 "alignment": {
								"horizontal": "center"
							}
						});
			var headings6 = stylesheet.createFormat({
							"font": {
								"size": 11,
								"color":"000000",
								"weight":'bold'
							},
						
							"fill": {
								"type": 'pattern',
								"patternType": 'solid',
								"fgColor": 'c5d9f1'
							},
							 "alignment": {
								"horizontal":"center"
							}
						});
				debugger;
				var mainData=[
							[
								{value: 'IT Capacity Filter Area................................................................................', metadata: {style: headings1.id}},
								'',
								'',
								'',
								'',
								'',
								'',
								'',
								''
							],
							[	
								'','','','','','','','',''
							],
							[	
								'','','','','','','','',''
							],
							[
								{value: 'ID', metadata: {style: headings6.id}},
								{value: 'By Providing Org', metadata: {style: headings6.id}},
								{value: 'By Providing Org Detail', metadata: {style: headings6.id}},
								{value: 'By Project Name (include "Org Level" to see values)', metadata: {style: headings6.id}},
								{value: 'By Work ID', metadata: {style: headings6.id}},
								{value: 'By For ICP', metadata: {style: headings6.id}},
								{value: 'By Labor Type', metadata: {style: headings6.id}},
								{value: 'By Labor Type', metadata: {style: headings6.id}},
								{value: 'Count', metadata: {style: headings6.id}}
							],
						];
			var columnWidth=[{width: 8},{width: 35},{width: 35},{width: 35},{width: 10},{width: 20},{width: 15},{width: 15},{width: 10}];
			for(var i=0,j=0;i<localData.length;i++){
				j=j+9;
				sheet1.mergeCells(cellsDefination[(j).toString()]+1,cellsDefination[(j+8).toString()]+1);
				sheet1.mergeCells(cellsDefination[(j).toString()]+2,cellsDefination[(j+8)].toString()+2);
				mainData[0].push({value:localData[i][0], metadata: {style: headings2.id}},'','','','','','','','');
				mainData[1].push({value:'Filtered Grand Total', metadata: {style: headings5.id}},'','','','','','','','');
				mainData[2].push({value:'=SUBTOTAL(109,'+cellsDefination[(j).toString()]+'5:'+cellsDefination[(j).toString()]+(localData2.length+5)+')', metadata: {style: headings3.id,type:'formula'}});
				mainData[2].push({value:'=IF('+cellsDefination[(j).toString()]+'3=0,0,('+cellsDefination[(j+2).toString()]+'3/'+cellsDefination[(j).toString()]+'3))', metadata: {style: headings3.id,type:'formula'}});
				mainData[2].push({value:'=SUBTOTAL(109,'+cellsDefination[(j+2).toString()]+'5:'+cellsDefination[(j+2).toString()]+(localData2.length+5)+')', metadata: {style: headings3.id,type:'formula'}});
				mainData[2].push({value:'=SUBTOTAL(109,'+cellsDefination[(j+3).toString()]+'5:'+cellsDefination[(j+3).toString()]+(localData2.length+5)+')', metadata: {style: headings3.id,type:'formula'}});
				mainData[2].push({value:'=SUBTOTAL(109,'+cellsDefination[(j+4).toString()]+'5:'+cellsDefination[(j+4).toString()]+(localData2.length+5)+')', metadata: {style: headings3.id,type:'formula'}});
				mainData[2].push({value:'=SUBTOTAL(109,'+cellsDefination[(j+5).toString()]+'5:'+cellsDefination[(j+5).toString()]+(localData2.length+5)+')', metadata: {style: headings3.id,type:'formula'}});
				mainData[2].push({value:'', metadata: {style: headings3.id}});
				mainData[2].push({value:'=SUBTOTAL(109,'+cellsDefination[(j+6).toString()]+'5:'+cellsDefination[(j+6).toString()]+(localData2.length+5)+')', metadata: {style: headings3.id,type:'formula'}});
				mainData[2].push({value:'=SUBTOTAL(109,'+cellsDefination[(j+7).toString()]+'5:'+cellsDefination[(j+7).toString()]+(localData2.length+5)+')', metadata: {style: headings3.id,type:'formula'}});
				mainData[3].push({value:'Org Capacity', metadata: {style: headings4.id}},
								{value:'% of Remaining Capacity', metadata: {style: headings4.id}},
								{value:'Org Remaining Capacity', metadata: {style: headings4.id}},
								{value:'Org Standard Capacity', metadata: {style: headings4.id}},
								{value:'Org Out of Port Demand', metadata: {style: headings4.id}},
								{value:'Org Demand', metadata: {style: headings4.id}},
								{value:'Org Demand Filtered', metadata: {style: headings4.id}},
								{value:'Project Demand', metadata: {style: headings4.id}},
								{value:'Project Demand Detail', metadata: {style: headings4.id}});
				columnWidth.push({width: 15});
			}
			
			mainData[3].push(	{value: 'Project Manager', metadata: {style: headings6.id}},
								{value: 'Project Status', metadata: {style: headings6.id}},
								{value: 'ITPortfolio Status', metadata: {style: headings6.id}},
								{value: 'PLT Name', metadata: {style: headings6.id}},
								{value: 'PLT-Priority', metadata: {style: headings6.id}},
								{value: 'Planned Production Date', metadata: {style: headings6.id}},
								{value: 'Project Type', metadata: {style: headings6.id}},
								{value: 'LRS Roadmap', metadata: {style: headings6.id}},
								{value: 'Work Type', metadata: {style: headings6.id}},
								{value: 'Investment Approval', metadata: {style: headings6.id}},
								{value: 'Project Classification', metadata: {style: headings6.id}}
							);
			var constant1=4,check=1;
			for(var i=0,a=mainData.length,count=5;i<localData2.length;i++,a++,count++){
				if(check || localData2[i][1]==localData2[i-1][1]){
					check=0;
					mainData[a]=[];
					for(var j=0;j<localData2[i].length;j++){
						if(j<10 || j>=localData2[i].length-10){
							if(j!=1)
								mainData[a].push(localData2[i][j])
						}
						else{
							mainData[a].push(localData2[i][j]);
							mainData[a].push({value:'=IF('+cellsDefination[(j-1).toString()]+count+'=0,0,('+cellsDefination[(j+1).toString()]+count+'/'+cellsDefination[(j-1).toString()]+count+'))', metadata: {type:'formula'}});
							mainData[a].push({value:'=IF(AND($'+cellsDefination[(j-6).toString()]+count+'="Org Level",ABS('+cellsDefination[(j+4).toString()]+count+'-'+cellsDefination[(j+5).toString()]+count+')<0.3),'+cellsDefination[(j-1).toString()]+count+'-'+cellsDefination[(j+2).toString()]+count+'-'+cellsDefination[(j+3).toString()]+count+'-'+cellsDefination[(j+4).toString()]+count+',IF('+cellsDefination[(j+5).toString()]+count+'=0,'+cellsDefination[(j-1).toString()]+count+'-'+cellsDefination[(j+2).toString()]+count+'-'+cellsDefination[(j+3).toString()]+count+'-'+cellsDefination[(j+4).toString()]+count+','+cellsDefination[(j-1).toString()]+count+'-'+cellsDefination[(j+2).toString()]+count+'-'+cellsDefination[(j+3).toString()]+count+'-'+cellsDefination[(j+5).toString()]+count+'))', metadata: {type:'formula'}});
							mainData[a].push(localData2[i][j+1]);
							mainData[a].push(localData2[i][j+2]);
							mainData[a].push(localData2[i][j+3]);
							mainData[a].push({value:'', metadata: {type:'formula'}});
							mainData[a].push(localData2[i][j+4]);
							mainData[a].push(localData2[i][j+5]);
							j=j+5;
						}
					}
				}
				else{
				if(i<localData2.length)
					{
						i--;
						a--;
						count--;
					}
					for(var l=15;l<mainData[constant1].length;l=l+9){
						mainData[constant1][l]={value:'=SUBTOTAL(109,'+cellsDefination[(l+1).toString()]+constant1+':'+cellsDefination[(l+1).toString()]+(mainData.length-1)+')', metadata: {type:'formula'}}
					}
					constant1=mainData.length;
					check=1;
				}
			}
			
					console.log('Processed');
					 sheet1.setColumns(columnWidth);
					sheet1.setData(mainData);
					workbook.worksheets=[];
					workbook.sharedStrings.stringArray=[];
					workbook.sharedStrings.strings={};
					workbook.addWorksheet(sheet1);
					console.log('Ab run hogi');
					ExcelBuilder.Builder.createFile(workbook, {
							type: "blob"
						}).then(function(dataTemp) {
							saveAs(new Blob([dataTemp], {
								type: "base64"
							}), "ICP.xlsx");
						});
		};