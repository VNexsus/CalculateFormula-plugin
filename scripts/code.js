/**
 *
 * (c) Copyright VNexsus 2025
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */

(function (window, undefined) {

    var theme = "light";
	var initialCell = null;
    var cellAddress = null;
    var processor = null;
	var calculateBtn = null, restartBtn = null, stepInBtn = null, stepOutBtn = null, closeBtn = null;
	var baseurl = document.location.protocol == 'file:' ? 'file:///' + document.location.pathname.substring(0, document.location.pathname.lastIndexOf("/")).substring(1) + '/' : document.location.href.substring(0, document.location.href.lastIndexOf("/")+1);
	var stepping = false;

    window.Asc.plugin.init = function () {
        document.body.classList.add(window.Asc.plugin.getEditorTheme());
        window.Asc.plugin.resizeWindow(600, 250, 600, 250, 800, 450);
        calculateBtn = byId('calculate');
        restartBtn = byId('restart');
        stepInBtn = byId('stepIn');
        stepOutBtn = byId('stepOut');
        closeBtn = byId('close');
        restartBtn.onclick = function () { window.restart() };
        calculateBtn.onclick = function () { window.calculate() };
        stepInBtn.onclick = function () { window.stepIn() };
        stepOutBtn.onclick = function () { window.stepOut() };
        closeBtn.onclick = function () { window.Asc.plugin.button(0); };
		var processorOptions = {
			outputElementId: 			'processed',
			formulaSeparators: 			parent.AscCommon.FormulaSeparators || { digitSeparator: '.', functionArgumentSeparator: ',' },
			cBoolLocal: 				parent.AscCommon.cBoolLocal || { t: 'TRUE', f: 'FALSE' },
			cFormulaFunctionToLocale: 	parent.AscCommonExcel.cFormulaFunctionToLocale || {},
			formulaRangeBorderColor: 	parent.AscCommonExcel.c_oAscFormulaRangeBorderColor || [{ r: 0, g: 0, b: 0, a: 1 }],
			CellFormat: 				parent.AscCommon.CellFormat,
			callback: 					window.command,
			ErrorToLocale:				function(value){ 
											var key = Object.keys(parent.AscCommon.cErrorOrigin).find(key => { return parent.AscCommon.cErrorOrigin[key] === value });
											return parent.AscCommon.cErrorLocal[key].replace('\\','') || parent.AscCommon.cErrorOrigin[key]
										}
		};
        processor = new ExpressionProcessor(processorOptions);
        window.processCell();
		parent.Asc.editor.controller.view.model.handlers.add('asc_onSelectionChanged', window.processCell);
    }
	
	window.calculate = function() {
		if (!processor || !processor.canCalculate())
			return;
		processor.calculateNextStep();
		window.refreshButtons();
	}
	
	window.restart = function(){
		if (!processor || !processor.canRevert())
			return;
		processor.revert()
		window.refreshButtons();
	}
	
	window.stepIn = function() {
		if (!processor || !processor.canCalculate())
			return;
		stepping = true;
		var nextNode = processor.canCalculate();
		if(nextNode){
			var sheetName = null;
			var cellAddress = null;
			if(nextNode.type == 'CELL') {
				sheetName = nextNode.token.ws.sName;
				cellAddress = nextNode.token.value;
			}
			else if(nextNode.type == 'NAME') {
				var linked = nextNode.token.Calculate();
				sheetName = linked.ws.sName;
				cellAddress = linked.value;
			}
			if(sheetName && cellAddress) {
				var sheet = parent.Asc.editor.GetSheet(sheetName);
				var cell = sheet.GetRange(cellAddress);
				formula = cell.GetFormula();
				value = cell.GetValue();
				if(formula == value)
					formula = "'" + value;
				window.parseFormula(formula, cell);
			}
		}
	}
	
	window.stepOut = function() {
		if (!processor || !processor.canStepOut())
			return;
		stepping = true;
		processor.stepOut();
		window.refreshButtons();
		window.drawSelection(processor.parser.tokens, processor.parser.cell);
		window.updateNote();
	}
	
	window.refreshButtons = function() {
		if(!processor)
			return;
        var nextNode = processor.canCalculate();
		if(!nextNode) {
			calculateBtn.disabled = true;
			calculateBtn.style.display = '';
			restartBtn.disabled = true;
			restartBtn.style.display = 'none';
			stepInBtn.disabled = true;
			stepOutBtn.disabled = !processor.canStepOut();
		}
		if(!nextNode && processor.canRevert()){
			calculateBtn.disabled = true;
			calculateBtn.style.display = 'none';
			restartBtn.style.display = '';
			restartBtn.disabled = false;
			stepOutBtn.disabled = !processor.canStepOut();
		}
		if(nextNode) {
			calculateBtn.disabled = false;
			calculateBtn.style.display = '';
			restartBtn.disabled = true;
			restartBtn.style.display = 'none';
			stepInBtn.disabled = true;
			stepOutBtn.disabled = true;
		}
		if(processor.canStepIn()) {
			stepInBtn.disabled = false;
		}
		if(processor.canStepOut()) {
			stepOutBtn.disabled = false;
		}
	}
	
    window.processCell = function () {
		if(stepping){
			stepping = false;
			return;
		}
		var formula = null;
        var value = null;
		var sheet = parent.Asc.editor.GetActiveSheet();
		if(!initialCell) {
			initialCell = sheet.GetActiveCell();
		}
		else {
			var cell = sheet.GetActiveCell()
			if(cell == initialCell)
				return;
			initialCell = sheet.GetActiveCell();
			processor.clear();
		}
		formula = initialCell.GetFormula();
		value = initialCell.GetValue();
		if(formula == value)
			formula = "'" + value;
        window.parseFormula(formula);
    }

    window.parseFormula = function (formula, cell = initialCell) {
		sheet = cell.GetWorksheet();
        var _parseResult = new parent.AscCommonExcel.ParseResult([], []);
        var U = new parent.AscCommonExcel.CCellWithFormula(sheet.worksheet, cell.range.bbox.r1, cell.range.bbox.c1);
		if(formula.startsWith('=')){
			formula = formula.substr(1);
			var formulaParser = new parent.AscCommonExcel.parserFormula(formula, U, sheet.worksheet);
			formulaParser.outStack.push = function(el){
				if(el && el.type && el.type == 11)
					_parseResult.elems.push(el);
			}
			formulaParser.parse(false, false, _parseResult, true);
		}
		else {
			formula = formula.substr(1);
			format = cell.GetNumberFormat();
			var f = new parent.AscCommon.CellFormat(format);
			var z = f.format(formula);
			var t = '';
			for(var i = 0; i < z.length; i++)
				if(!z[i].hasOwnProperty('format'))
					t+=z[i].text
			var el = new parent.AscCommonExcel.cBaseType(formula.length > 0 ? t : '');
			el.type = -1; 
			_parseResult.elems.push(el);
		}
        const result = processor.process(_parseResult.elems, cell);
		window.refreshButtons();
		window.drawSelection(_parseResult.elems, cell);
		window.updateNote();
    }

	window.drawSelection = function(tokens, cell = initialCell) {
		parent.Asc.editor.controller.view.wsViews.forEach( v => { if(v){v.cleanSelection();v.oOtherRanges = null;v.updateSelection()}});
		if(cell && tokens && tokens.length > 0) {
			cell.Worksheet.SetActive();
			cell.Select();
			parent.Asc.editor.controller.view.resize();
			parent.SSE.controllers.Statusbar.selectTab(cell.Worksheet.Index);
			var seenCells = new Map();
			var selection = new parent.AscCommonExcel.SelectionRange(cell ? cell.GetWorksheet() : initialCell.GetWorksheet());
			selection.ranges.length = 0;
			tokens.forEach(el => {
				if (el.type === 6 || el.type === 5 || el.type === 12 || el.type === 13 || el.type === 10 || el.type === 15) {
					if(el.type == 10 || el.type == 15) el = el.Calculate();
					var ref = (el.ws ? el.ws.sName : el.wsFrom ? el.wsFrom.sName : '' ) + '!' + el.value;
					var rng = el.getBBox0().clone();
					if(!seenCells.get(ref)) {
						seenCells.set(ref, 1);
						if((el.ws && el.ws.sName != cell.Worksheet.worksheet.sName) || (el.wsFrom && el.wsFrom.sName != cell.Worksheet.worksheet.sName))
							rng.c1 = rng.c1 ? -rng.c1 : -1, rng.c2 = rng.c2 ? -rng.c2 : -1, rng.r1 = rng.r1 ? -rng.r1 : -1, rng.r2 = rng.r2 ? -rng.r2 : -1;
						rng.isName = true;
						selection.ranges.push(rng);
					}
				}
			});
			var view = parent.Asc.editor.controller.view.wsViews.find(el => el && el.model === cell.Worksheet.worksheet);
			if(view) {
				view.oOtherRanges = selection;
				view._drawSelection();
			}
		}
		if(stepping){
			stepping = false;
		}
	}
	
	window.updateNote = function() {
		if(processor && processor.parser && processor.parser.AST) {
			if(processor.parser.AST.token && processor.parser.AST.token.type == -1 && processor.parser.AST.token.value == '')
				document.querySelector('.note').innerText = 'Выбранная для вычисления ячейка пуста.';
			else if(processor.parser.AST.token && processor.parser.AST.token.type == -1)
				document.querySelector('.note').innerText = 'Выбранная для вычисления ячейка содержит константу.';
			else
				document.querySelector('.note').innerText = 'Для вычисления подчеркнутого красным выражения нажмите кнопку "Вычислить". Последний полученный результат выделяется курсивом.';
		}
	}

	window.command = function(message) {
		if(!message || !message.command)
			return;
		switch(message.command) {
			case 'functionHelp':
				if(message.name) {
					var url = parent.document.location.href;
					url = url.substr(0, url.indexOf('?'))
					url = url.substr(0, url.lastIndexOf('/') + 1) + 'resources/help/ru/';
					var e = new parent.Common.UI.Window({
						title: 		'Справка по функции ' + (message.nameLocal ? message.nameLocal : message.name),
						cls: 		"helpdialog",
						callback: 	function(){e.close()}
					});
					e.setElement(window.document);
					e.options.tpl = `<div class="info-box" style="height:100%;overflow: auto;"><div class="text" style="position:relative;margin-right: 10px;" tabindex="-1"></div></div><link rel='stylesheet' href='`+ baseurl +`resources/css/help.css'>`;
					e.on("render:after", function(){
						window.Asc.plugin.loadModule(url + 'Functions/' + message.name.toLowerCase().replaceAll('.','-') + '.htm', function(content){
							if(content){
								var parser = new DOMParser();
								var doc = parser.parseFromString(content,"text/html");
								doc.querySelector('.search-field').remove();
								doc.querySelectorAll('p').forEach( p => {p.style.textIndent = p.style.textIndent ? '20px' : ''});
								doc.querySelectorAll('table').forEach( p => {p.style.width = '100%'});
								doc.querySelectorAll('img').forEach( img => {img.src = url + 'images/' + img.src.substr(img.src.lastIndexOf('/'))});
								var help = doc.querySelector('.mainpart');
								e.$window[0].querySelector('.text').innerHTML = help.innerHTML;
								e.$window[0].querySelector('.text').querySelectorAll('a').forEach( a => { 
									var target = a.href.substring(a.href.lastIndexOf('/') + 1, a.href.lastIndexOf('.'));
									a.href = "#";
									a.removeAttribute("onclick");
									a.addEventListener('click', function(){ e.close(); window.command({command: 'functionHelp', name: target, nameLocal: a.innerText})});
								});
								e.$window[0].querySelector('.text').addEventListener('contextmenu', function(evt){ evt.returnValue = true; evt.stopPropagation(); evt.cancelBubble(); return true; })
							}
							else
								e.$window[0].querySelector('.text').innerHTML = 'Не удалось загрузить справку по этой функции';
							e.$window[0].setAttribute('tabindex', -1);
							e.$window[0].addEventListener('keydown', (evt) => {
								if(evt.keyCode == 27) e.close()
								if((evt.ctrlKey || evt.metaKey) && evt.keyCode == 65) {
									event.preventDefault();
									const targetDiv = e.$window[0].querySelector('.text');
									const range = document.createRange();
									range.selectNodeContents(targetDiv);
									const selection = parent.window.getSelection();
									selection.removeAllRanges();
									selection.addRange(range);
								}
							});
							e.$window[0].querySelector('.text').focus();
						});
					})
					.on('close', function() {var mask = parent.document.querySelector('.modals-mask'); mask && mask.remove();});
					e.render();
					e.setSize(600,400);
					e.$window[0].style.left = (parent.window.innerWidth - e.$window.width())/2 + "px";
					e.$window[0].style.top = (parent.window.innerHeight - e.$window.height())/2 - 4 + "px"; 
					e.show();
				}
				break;
		}
	}

    window.Asc.plugin.button = function (id) {
		parent.Asc.editor.controller.view.model.handlers.remove('asc_onSelectionChanged', window.processCell);
		initialCell.Select();
		drawSelection();
		parent.Asc.editor.controller.view.resize();
		parent.focus();
        this.executeCommand("close", "");
    };

    window.Asc.plugin.onThemeChanged = function (theme) {
        window.Asc.plugin.onThemeChangedBase(theme);
        document.body.classList.remove("theme-dark", "theme-light");
        document.body.classList.add(window.Asc.plugin.getEditorTheme());
    }

    window.Asc.plugin.getEditorTheme = function () {
        if (window.localStorage && window.localStorage.getItem("ui-theme")) {
            var x = JSON.parse(window.localStorage.getItem("ui-theme"));
            theme = x.type;
            return 'theme-' + x.type;
        }
        theme = 'light';
        return "theme-light";
    }

    function byId(id) { return document.getElementById(id); }

})(window, undefined);