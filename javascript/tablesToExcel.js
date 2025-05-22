//Class to convert an array of HTML table elements into sheets in an XML-based Excel document.
//This does not retain styling.
//Behavior when dealing with malformed HTML tables is undefined.
//@param {Array<HTMLTableElement>} tableElems
//@param {Array<string>} sheetnames
//@param {object} documentProperties            Excel document metadata (e.g., Title, Subject, Author, Keywords, Description, Created)
//@return {string}                              XML
function ExcelWorkbook(tableElems, sheetnames = [], documentProperties = {}){
	
	'use strict';

	let xml,
		href;

	//constructor
	(function (){
		
		let sheetNum = 1;
		
		xml =	'<?xml version="1.0" encoding="UTF-8"?><?mso-application progid="Excel.Sheet"?>'+
				'<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"'+
						' xmlns:o="urn:schemas-microsoft-com:office:office"'+
						' xmlns:x="urn:schemas-microsoft-com:office:excel"'+
						' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"'+
						' xmlns:html="http://www.w3.org/TR/REC-html40">'+
					'<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">';
		for(let prop of documentProperties.getOwnPropertyNames()){
			if(/^xml|^[^a-z_]|[^a-z0-9_.-]/i.test(prop)) continue;	//invalid XML element name
			if(documentProperties[prop] instanceof Date) documentProperties[prop] = documentProperties[prop].toISOString();
			xml +=		'<'+prop+'>'+escapeXML(documentProperties[prop])+'</'+prop+'>';
		}
		xml +=		'</DocumentProperties>'+
					'<ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">'+
						'<ProtectStructure>False</ProtectStructure>'+
						'<ProtectWindows>False</ProtectWindows>'+
					'</ExcelWorkbook>'+
					'<Styles>'+
						'<Style ss:ID="Default" ss:Name="Normal">'+
							'<Alignment ss:Vertical="Top" ss:Horizontal="Left" ss:WrapText="0"/>'+
						'</Style>'+
						'<Style ss:ID="th">'+
							'<Alignment ss:Vertical="Bottom"/>'+
							'<Font ss:Bold="1"/>'+
						'</Style>'+
					'</Styles>';

		for(let i=0; i<tableElems.length; i++){
			xml += addSheet(tableElems[i], sheetnames[i]);
		}

		xml +=	'</Workbook>';

		function escapeXML(unsafe){
			return unsafe.replace(/[<>&"'']/g, function (char){
				switch(char){
					case '<': return '&lt;';
					case '>': return '&gt;';
					case '&': return '&amp;';
					case '"': return '&quot;';
					case '\'': return '&apos;';
				}
			});
		}

		function addSheet(tableElem, sheetname){

			let rowCount = 0, columnCount = 0,
				headerRowCount = 0, headerComplete,
				rowspans = [];

			xml +=	'<Worksheet ss:Name="'+(sheetname || sheetname === 0 ? sheetname : 'Sheet'+(sheetNum++))+'">'+
						'<Table>';

			for(let elem of tableElem.children){
				if(elem.tagName === 'THEAD'){
					for(let row of elem.children){
						if(!headerComplete) headerRowCount++;
						addRow(row, 'th');
					}
				}
				else if(elem.tagName === 'TBODY'){
					headerComplete = true;
					for(let row of elem.children){
						addRow(row);
					}
				}
				else if(elem.tagName === 'TR'){
					if(elem.children[0].tagName === 'TH'){
						if(!headerComplete) headerRowCount++;
						addRow(elem, 'th');
					}
					else{
						headerComplete = true;
						addRow(elem);
					}
				}
			}

			xml +=		'</Table>'+
						'<WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">'+
							'<Selected/>';
			if(headerRowCount){ xml +=
							'<FreezePanes/>'+
							'<FrozenNoSplit/>'+
							'<SplitHorizontal>'+headerRowCount+'</SplitHorizontal>'+
							'<TopRowBottomPane>1</TopRowBottomPane>'+
							'<ActivePane>2</ActivePane>'+
							'<Panes>'+
								'<Pane><Number>3</Number></Pane>'+
								'<Pane><Number>2</Number></Pane>'+
							'</Panes>';
			}
			xml +=			'<ProtectObjects>False</ProtectObjects>'+
							'<ProtectScenarios>False</ProtectScenarios>'+
						'</WorksheetOptions>';
			if(headerRowCount){ xml +=
						'<AutoFilter xmlns="urn:schemas-microsoft-com:office:excel" x:Range="R'+headerRowCount+'C1:R'+rowCount+'C'+columnCount+'"/>';
			}

			xml +=	'</Worksheet>';

			function addRow(row, styleID){

				let cells = row.children,
					col=0;

				xml += '<Row>';

				if(!cells.length){
					//no cell elements, so they must all be part of existing row spans
					rowspans.forEach( (span)=>span-- );
				}
				else{
					for(let c=0; c<cells.length; col++){

						if(rowspans[col]){
							//the cell for this column is part of an existing row span
							rowspans[col]--;
							continue;
						}

						let cell = cells[c++],
							data = cell.textContent,
							dataType = /^-?\d+(\.\d+)?$/.test(data) ? 'Number' : 'String',
							colspan = cell.colSpan-1,	//number of additional columns to merge with
							rowspan = cell.rowSpan-1;	//number of additional rows to merge with

						if(rowspan) rowspans[col] += rowspan;

						xml += '<Cell'+
							(colspan ? ' ss:MergeAcross="'+colspan+'"' : '' )+
							(rowspan ? ' ss:MergeDown="'+rowspan+'"' : '')+
							(styleID ? ' ss:StyleID="'+styleID+'"' : '')+
							'><Data ss:Type="'+dataType+'">'+data+'</Data></Cell>';

						if(colspan){
							//update row span counts for merged cells
							for(let i=1; i<=colspan; i++){
								if(rowspans[col+i]) rowspans[col+i]--;
							}
							col += colspan;
						}

					}
				}

				xml += '</Row>';

				rowCount++;
				if(!columnCount) columnCount = col;

			}

		}
		
	})();
	
	//if called with the `new` operator
	Object.defineProperties(this, {
		xml: {
			get: () => xml
		},
		href: {
			get: () => { return href || (href = URL.createObjectURL( new Blob([xml], {type:'application/vnd.ms-excel'}) )); }
		},
		revoke: {
			value: () => { href = URL.revokeObjectURL(href); }
		}
	});

	//if called without the `new` operator
	return xml;

}	
