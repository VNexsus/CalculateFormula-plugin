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

const NodeType = {
    DEFAULT: 'DEFAULT',
	NUMBER: 'NUMBER',
    STRING: 'STRING',
    BOOLEAN: 'BOOLEAN',
    BINARY_OPERATION: 'BINARY_OPERATION',
    UNARY_OPERATION: 'UNARY_OPERATION',
    PARENTHESIS: 'PARENTHESIS',
    FUNCTION_CALL: 'FUNCTION_CALL',
    CELL: 'CELL',
    CELLRANGE: 'CELLRANGE',
	NAME: 'NAME',
	ERROR: 'ERROR',
	ARRAY: 'ARRAY'
};

var counter = 10000;

// Классы узлов AST
class ASTNode {
    constructor(type) {
        this.type = type;
        this.el = null;
        this.parent = null;
        this.id = counter++;
    }
}

class DefaultNode extends ASTNode {
    constructor(token, options = {}) {
        super(NodeType.DEFAULT);
        this.token = token;
        this.value = token.value;
		this.options = options;
    }
	
    toHtml() {
        this.el = document.createElement('span');
        this.el.setAttribute('id', 'node-' + this.id)
        this.el.classList.add('default');
        this.el.textContent = this.value;
        return this.el;
    }
}

class EmptyNode extends ASTNode {
    constructor(token, options = {}) {
        super(NodeType.DEFAULT);
        this.token = token;
        this.value = 0;
		this.options = options;
    }
	
    toHtml() {
        this.el = document.createElement('span');
        this.el.setAttribute('id', 'node-' + this.id)
        this.el.classList.add('empty');
        this.el.textContent = this.value;
        return this.el;
    }
}

class NumberNode extends ASTNode {
    constructor(token, options = {}) {
        super(NodeType.NUMBER);
        this.token = token;
        this.value = this.fixNumber(token.value);
		this.options = options;
    }
	
	fixNumber (num){
		var s = String(num);
		s = s.substr(0, s.indexOf('.') == -1 ? 15 : 16)
		return parseFloat(s);
	}

    toHtml() {
		this.el = document.createElement('span');
		this.el.setAttribute('id', 'node-' + this.id);
		this.el.classList.add('number');

		let text = null;

		const isRootResult = !this.parent;

		try {
			if (isRootResult && this.options.cell && this.options.CellFormat) {
				const format = this.options.cell.GetNumberFormat();
				const f = new this.options.CellFormat(format);
				const z = f.format(this.value);
				let t = '';
				for (let i = 0; i < z.length; i++) {
					if (!z[i].hasOwnProperty('format'))
						t += z[i].text;
				}
				text = t;
			}
		} 
		catch (e) {
			console.error('Ошибка форматирования числа:', e);
		}
		
		if (!text) {
			text = String(this.value).replace('.', this.options.formulaSeparators.digitSeparator);
		}

		this.el.textContent = text;

		this.el.addEventListener('click', (e) => {
			e.stopPropagation();
		});
		return this.el;
	}

}

class StringNode extends ASTNode {
    constructor(token, options = {}) {
        super(NodeType.STRING);
        this.token = token;
        this.value = token.value;
		this.options = options;
    }

    toHtml() {
        this.el = document.createElement('span');
        this.el.setAttribute('id', 'node-' + this.id)
        this.el.classList.add('string');
        this.el.textContent = '"' + this.value + '"';
        this.el.addEventListener('click', (e) => {
            e.stopPropagation();
        });
        return this.el;
    }
}

class BooleanNode extends ASTNode {
    constructor(token, options = {}) {
        super(NodeType.BOOLEAN);
        this.token = token;
        this.value = token.value;
		this.options = options;
    }

    toHtml() {
        this.el = document.createElement('span');
        this.el.setAttribute('id', 'node-' + this.id)
        this.el.classList.add('boolean');
        this.el.textContent = this.value === true
            ? this.options.cBoolLocal.t || 'TRUE'
            : this.options.cBoolLocal.f || 'FALSE';
        this.el.addEventListener('click', (e) => {
            e.stopPropagation();
        });
        return this.el;
    }
}

class ErrorNode extends ASTNode {
    constructor(token, options = {}) {
        super(NodeType.ERROR);
        this.token = token;
        this.value = token.value;
		this.options = options;
    }
	
	toHtml() {
        this.el = document.createElement('span');
        this.el.setAttribute('id', 'node-' + this.id);
        this.el.classList.add('error');
        this.el.textContent = this.options.ErrorToLocale(this.value);
        this.el.addEventListener('click', (e) => {
            e.stopPropagation();
        });
        return this.el;
    }
}

class CellNode extends ASTNode {
    constructor(token, options = {}, color) {
        super(NodeType.CELL);
        this.token = token;
		this.options = options;
		this.color = color;
		let value = token.value;
		
        if (token.type === 12 && token.ws && token.ws.sName) {
            let sheetName = token.ws.sName;

            if (sheetName.indexOf(' ') >= 0) {
                sheetName = "'" + sheetName + "'";
            }
            if (value && value.indexOf('!') < 0) {
                value = sheetName + '!' + value;
            }
        }

        this.value = value;
    }

    toHtml() {
        this.el = document.createElement('span');
        this.el.setAttribute('id', 'node-' + this.id)
        this.el.classList.add('cell', 'expression');

        const addr = document.createElement('span');
        addr.className = 'cell-address';
        addr.textContent = this.value;
		addr.style.color = this.color;
		this.el.appendChild(addr);
        return this.el;
    }
	
	calculate() {
		var result = this.token.getValue();
		let resultNode;

		switch (result.type) {
			case 0: resultNode = new NumberNode(result, this.options); break;
			case 1: resultNode = new StringNode(result, this.options); break;
			case 2: resultNode = new BooleanNode(result, this.options); break;
			case 3: resultNode = new ErrorNode(result, this.options); break;
			case 4: resultNode = new EmptyNode(result, this.options); break;
			default: resultNode = new NumberNode(result, this.options);
		}

		var el = document.getElementById('node-' + this.id);

		if (this.parent) {
			resultNode.parent = this.parent;
			switch (this.parent.type) {
				case NodeType.BINARY_OPERATION:
					if (this.parent.left == this) this.parent.left = resultNode;
					if (this.parent.right == this) this.parent.right = resultNode;
					break;
				case NodeType.PARENTHESIS:
					this.parent.expression = resultNode;
					break;
				case NodeType.FUNCTION_CALL:
					const argIndex = this.parent.arguments.indexOf(this);
					if (argIndex !== -1)
						this.parent.arguments[argIndex] = resultNode;
					break;
				case NodeType.UNARY_OPERATION:
					break;
			}
		}

		resultNode.toHtml();
		resultNode.el.classList.add('last');
		resultNode.el.style.color = this.color;
		el.replaceWith(resultNode.el);

		if (!this.parent) {
			return resultNode;
		}
	}
}

class CellRangeNode extends ASTNode {
    constructor(token, options = {}, color) {
        super(NodeType.CELLRANGE);
        this.token = token;
        this.options = options;
        this.color = color;

        let value = token.value;
        let sheetName = null;

        if (token.type === 13 && token.ws && token.ws.sName) {
            sheetName = token.ws.sName;
        } else {
            const fromName = token.wsFrom && token.wsFrom.sName;
            const toName   = token.wsTo   && token.wsTo.sName;

            sheetName = fromName || toName || null;
        }

        if (sheetName && value && value.indexOf('!') < 0) {
            if (sheetName.indexOf(' ') >= 0) {
                sheetName = "'" + sheetName + "'";
            }
            value = sheetName + '!' + value;
        }

        this.value = value;
    }

    toHtml() {
        this.el = document.createElement('span');
        this.el.setAttribute('id', 'node-' + this.id)
        this.el.classList.add('cell-range', 'expression');
        this.el.textContent = this.value;
        this.el.style.color = this.color;
        return this.el;
    }
}

class ArrayNode extends ASTNode {
    constructor(token, options = {}) {
        super(NodeType.ARRAY);
        this.token = token;
		this.options = options;
    }

    toHtml() {
        this.el = document.createElement('span');
        this.el.setAttribute('id', 'node-' + this.id)
        this.el.classList.add('array');
        this.el.textContent = this.token.toLocaleString(true);
		return this.el;
    }
}

class NamedRangeNode extends ASTNode {
    constructor(token, options = {}, color) {
        super(NodeType.NAME);
        this.token = token;
        this.value = token.value;
		this.options = options;
		this.color = color;
    }

    toHtml() {
        this.el = document.createElement('span');
        this.el.setAttribute('id', 'node-' + this.id)
        this.el.classList.add('named-range', 'expression');
        this.el.textContent = this.value;
		this.el.style.color = this.color;
		return this.el;
    }
	
	calculate() {
		var linkedCell = this.token.Calculate();
		var result = linkedCell.getValue();
		let resultNode;

		switch (result.type) {
			case 0: resultNode = new NumberNode(result, this.options); break;
			case 1: resultNode = new StringNode(result, this.options); break;
			case 2: resultNode = new BooleanNode(result, this.options); break;
			case 3: resultNode = new ErrorNode(result, this.options); break;
			case 4: resultNode = new EmptyNode(result, this.options); break;
			default: resultNode = new NumberNode(result, this.options);
		}

		var el = document.getElementById('node-' + this.id);

		if (this.parent) {
			resultNode.parent = this.parent;
			switch (this.parent.type) {
				case NodeType.BINARY_OPERATION:
					if (this.parent.left == this) this.parent.left = resultNode;
					if (this.parent.right == this) this.parent.right = resultNode;
					break;
				case NodeType.PARENTHESIS:
					this.parent.expression = resultNode;
					break;
				case NodeType.FUNCTION_CALL:
					const argIndex = this.parent.arguments.indexOf(this);
					if (argIndex !== -1)
						this.parent.arguments[argIndex] = resultNode;
					break;
				case NodeType.UNARY_OPERATION:
					break;
			}
		}

		resultNode.toHtml();
		resultNode.el.classList.add('last');
		resultNode.el.style.color = this.color;
		el.replaceWith(resultNode.el);

		if (!this.parent) {
			return resultNode;
		}
	}
}

function getTokenFromNode(node) {
    if (!node) return null;

    if (node.type === NodeType.UNARY_OPERATION) {
        const valueToken = getTokenFromNode(node.argument);
        try {
            return node.token.Calculate([valueToken]);
        } catch (e) {
            console.error('Ошибка унарного оператора:', e);
            return valueToken;
        }
    }

    if (node.type === NodeType.PARENTHESIS) {
        return getTokenFromNode(node.expression);
    }

    if (node.type === NodeType.NAME) {
		var result = node.token.Calculate();
		return result.getValue();
	}

    return node.token;
}


class BinaryOperationNode extends ASTNode {
    constructor(token, left, right, options = {}) {
        super(NodeType.BINARY_OPERATION);
        this.operator = token;
        this.left = left, this.left.parent = this;
        this.right = right, this.right.parent = this;
		this.options = options; 

        NodeList[this.left] ? NodeList[this.left].push(this)
            : NodeList[this.left] = [this];
        NodeList[this.right] ? NodeList[this.right].push(this)
            : NodeList[this.right] = [this];
    }

    toHtml() {
        this.el = document.createElement('span');
        this.el.setAttribute('id', 'node-' + this.id)
        this.el.classList.add('expression');

        const op = document.createElement('span');
        op.className = 'binary-operator';
        op.textContent = this.operator.name || this.operator.value;

        this.el.appendChild(this.left.toHtml());
        this.el.appendChild(op);
        this.el.appendChild(this.right.toHtml());

        this.el.addEventListener('click', (e) => {
            e.stopPropagation();
        });

        return this.el;
    }

    calculate() {
        try {
			const leftToken  = getTokenFromNode(this.left);
			const rightToken = getTokenFromNode(this.right);
			var result = this.operator.Calculate([leftToken, rightToken]);
            let resultNode;

            switch (result.type) {
                case 0: resultNode = new NumberNode(result, this.options); break;
                case 1: resultNode = new StringNode(result, this.options); break;
                case 2: resultNode = new BooleanNode(result, this.options); break;
				case 3: resultNode = new ErrorNode(result, this.options); break;
				case 4: resultNode = new EmptyNode(result, this.options); break;
				case 11: resultNode = new ArrayNode(result, this.options); break;
                default: resultNode = new NumberNode(result, this.options);
            }

			var el = document.getElementById('node-' + this.id);

			if (this.parent) {
				resultNode.parent = this.parent; 
				switch (this.parent.type) {
					case NodeType.BINARY_OPERATION:
						if (this.parent.left == this) this.parent.left = resultNode;
						if (this.parent.right == this) this.parent.right = resultNode;
						break;
					case NodeType.PARENTHESIS:
						this.parent.expression = resultNode;
						break;
					case NodeType.FUNCTION_CALL:
					const argIndex = this.parent.arguments.indexOf(this);
						if (argIndex !== -1)
							this.parent.arguments[argIndex] = resultNode;
						break;
					case NodeType.UNARY_OPERATION:
						break;
				}
			}


			resultNode.toHtml();
			resultNode.el.classList.add('last');
			el.replaceWith(resultNode.el);

			if (!this.parent) {
				return resultNode;
			}
		} 
		catch (error) {
				console.error('Ошибка вычисления:', error);
				return false;
		}
	}
}
		
class UnaryOperationNode extends ASTNode {
    constructor(token, argument, options = {}) {
        super(NodeType.UNARY_OPERATION);
        this.token = token;
		const nameOrValue =
            (token.value !== undefined && token.value !== null)
                ? token.value
                : token.name;
        this.operator = (token.name === 'un_minus')
            ? '-'
            : (token.name === 'un_plus' ? '+' : token.name);
        this.argument = argument;
		this.argument.parent = this;
		this.options = options; 
    }

    toHtml() {
        this.el = document.createElement('span');
        this.el.setAttribute('id', 'node-' + this.id)
        this.el.classList.add('expression', 'unary');

        const op = document.createElement('span');
        op.className = 'unary-operator';
        op.textContent = this.operator;

        const argHtml = this.argument.toHtml();

        if (this.operator === '-' || this.operator === '+') {
            this.el.appendChild(op);
            this.el.appendChild(argHtml);
        }
        else if (this.operator === '%') {
            this.el.appendChild(argHtml);
            this.el.appendChild(op);
        } else {
            this.el.appendChild(op);
            this.el.appendChild(argHtml);
        }

        this.el.addEventListener('click', (e) => {
            e.stopPropagation();
        });

        return this.el;
    }
	
	calculate() {
		try {
			const valueToken = getTokenFromNode(this.argument);
			const result = this.token.Calculate([valueToken]);
			let resultNode;

			switch (result.type) {
				case 0: resultNode = new NumberNode(result, this.options); break;
				case 1: resultNode = new StringNode(result, this.options); break;
				case 2: resultNode = new BooleanNode(result, this.options); break;
				case 3: resultNode = new ErrorNode(result, this.options); break;
				case 4: resultNode = new EmptyNode(result, this.options); break;
				default: resultNode = new NumberNode(result, this.options);
			}

			const el = document.getElementById('node-' + this.id);

			if (this.parent) {
				resultNode.parent = this.parent;
				switch (this.parent.type) {
					case NodeType.BINARY_OPERATION:
						if (this.parent.left === this)  this.parent.left  = resultNode;
						if (this.parent.right === this) this.parent.right = resultNode;
						break;

					case NodeType.PARENTHESIS:
						this.parent.expression = resultNode;
						break;

					case NodeType.FUNCTION_CALL: {
						const idx = this.parent.arguments.indexOf(this);
						if (idx !== -1)
							this.parent.arguments[idx] = resultNode;
						break;
					}

					case NodeType.UNARY_OPERATION:
						this.parent.argument = resultNode;
						break;
				}
			}

			resultNode.toHtml();
			resultNode.el.classList.add('last');
			el.replaceWith(resultNode.el);

			if (!this.parent) {
				return resultNode;
			}
		} 
		catch (e) {
			console.error('Ошибка вычисления унарной операции:', e);
			return false;
		}
	}
}

class ParenthesisNode extends ASTNode {
    constructor(expression) {
        super(NodeType.PARENTHESIS);
        this.expression = expression;
        this.expression.parent = this;
    }

    toHtml() {
        this.el = document.createElement('span');
        this.el.setAttribute('id', 'node-' + this.id)
        this.el.classList.add('expression', 'parenthesis');

        const open = document.createElement('span');
        open.className = 'paren-open';
        open.textContent = '(';

        const content = this.expression.toHtml();

        const close = document.createElement('span');
        close.className = 'paren-close';
        close.textContent = ')';

        this.el.appendChild(open);
        this.el.appendChild(content);
        this.el.appendChild(close);

        this.el.addEventListener('click', e => e.stopPropagation());

        return this.el;
    }

    calculate() {
        this.expression.el.classList.add('last');

        if (this.parent) {
			this.expression.parent = this.parent;
            switch (this.parent.type) {
                case NodeType.BINARY_OPERATION:
                    var operand = this.parent[this.parent.left == this ? 'left' : 'right'];
                    var el = document.getElementById('node-' + operand.id);
                    el.replaceWith(this.expression.el);
                    this.parent[this.parent.left == this ? 'left' : 'right'] = this.expression;
                    break;

                case NodeType.PARENTHESIS:
                    var el = document.getElementById('node-' + this.id);
                    el.replaceWith(this.expression.el);
                    this.parent.expression = this.expression;
                    break;

                case NodeType.FUNCTION_CALL:
                    var el = document.getElementById('node-' + this.id);
					var idx = this.parent.arguments.indexOf(this);
					if (idx !== -1) {
						el.replaceWith(this.expression.el);
						this.parent.arguments[idx] = this.expression;
					}
                    break;

                case NodeType.UNARY_OPERATION:
                    var el = document.getElementById('node-' + this.id);
                    el.replaceWith(this.expression.el);
					this.parent.argument = this.expression;
                    break;
            }
        } 
		else {
			var el = document.getElementById('node-' + this.id);
			el.replaceWith(this.expression.el);
            return this.expression;
        }
    }
}

class FunctionCallNode extends ASTNode {
    constructor(token, args, options = {}) {
        super(NodeType.FUNCTION_CALL);
        this.token = token;
        this.name = token.name;
        this.arguments = args || [];
        this.arguments.forEach(arg => { arg.parent = this });
		this.options = options;
    }
	
    toHtml() {
        this.el = document.createElement('span');
        this.el.setAttribute('id', 'node-' + this.id)
        this.el.classList.add('function-call', 'expression');
        this.el.addEventListener('click', (e) => {
            e.stopPropagation();
        });

        const fn = document.createElement('span');
        fn.className = 'function-name';
        fn.textContent = (this.options.cFormulaFunctionToLocale && this.options.cFormulaFunctionToLocale[this.name]) || this.name;
        this.el.appendChild(fn);
		fn.addEventListener('click', () => openHelp(this.name, this.options));
		function openHelp(name, options){
			if(name) {
				options.callback({
					command: 'functionHelp', 
					name: name,
					nameLocal: options.cFormulaFunctionToLocale && options.cFormulaFunctionToLocale[name]
				})
			}
		}

        const open = document.createElement('span');
        open.className = 'paren-open';
        open.textContent = '(';
        this.el.appendChild(open);

        this.arguments.forEach((arg, idx) => {
            const argWrap = document.createElement('span');
            argWrap.className = 'function-arg';
            argWrap.appendChild(arg.toHtml());
            this.el.appendChild(argWrap);

            if (idx < this.arguments.length - 1) {
                const sep = document.createElement('span');
                sep.className = 'arg-separator';
                sep.textContent = this.options.formulaSeparators.functionArgumentSeparator;
                this.el.appendChild(sep);
            }
        });

        const close = document.createElement('span');
        close.className = 'paren-close';
        close.textContent = ')';
        this.el.appendChild(close);

        return this.el;
    }

	calculate() {
		var args = [];
		this.arguments.forEach(arg => args.push(getTokenFromNode(arg)));

		let result;
		if (['INDIRECT','SHEET','SHEETS'].indexOf(this.token.name) >= 0)
			result = this.token.Calculate(args, null, null, this.options.cell.Worksheet.worksheet);
		else
			result = this.token.Calculate(args);

		if (result.type == 6 || result.type == 12)
			result = result.getValue();

		let resultNode;
		switch (result.type) {
			case 0: resultNode = new NumberNode(result, this.options); break;
			case 1: resultNode = new StringNode(result, this.options); break;
			case 2: resultNode = new BooleanNode(result, this.options); break;
			case 3: resultNode = new ErrorNode(result, this.options); break;
			case 4: resultNode = new EmptyNode(result, this.options); break;
			case 11: resultNode = new ArrayNode(result, this.options); break;
			default: resultNode = new NumberNode(result, this.options);
		}

		const el = document.getElementById('node-' + this.id);

		if (this.parent) {
			resultNode.parent = this.parent;
			switch (this.parent.type) {
				case NodeType.BINARY_OPERATION:
					if (this.parent.left == this)  this.parent.left  = resultNode;
					if (this.parent.right == this) this.parent.right = resultNode;
					break;

				case NodeType.PARENTHESIS:
					this.parent.expression = resultNode;
					break;

				case NodeType.FUNCTION_CALL: {
					const argIndex = this.parent.arguments.indexOf(this);
					if (argIndex !== -1)
						this.parent.arguments[argIndex] = resultNode;
					break;
				}

				case NodeType.UNARY_OPERATION:
					break;
			}
		}

		resultNode.toHtml();
		resultNode.el.classList.add('last');
		el.replaceWith(resultNode.el);

		if (!this.parent)
			return resultNode;
	}
}

// Парсер AST
class ASTParser {
    constructor(cell, options = {}) {
		this.operators = {
            '^': { precedence: 4, associativity: 'right' },
            '*': { precedence: 3, associativity: 'left' },
            '/': { precedence: 3, associativity: 'left' },
            '+': { precedence: 2, associativity: 'left' },
            '-': { precedence: 2, associativity: 'left' },
            '&': { precedence: 1, associativity: 'left' },
            '=': { precedence: 0, associativity: 'left' },
            '>': { precedence: 0, associativity: 'left' },
            '<': { precedence: 0, associativity: 'left' },
            '>=': { precedence: 0, associativity: 'left' },
            '<=': { precedence: 0, associativity: 'left' },
            '<>': { precedence: 0, associativity: 'left' }
        };
        this.currentTokenIndex = 0;
        this.tokens = [];
        this.AST = null;
        this.highestPriorityNode = null;
		this.options = Object.assign({}, options, { cell });
		this.cell = cell;
		this.seenCells = new Map();
		this.id = counter++;
    }
	
	isPercentToken(token) {
        if (!token || token.type !== 9)
            return false;

        const v = (token.value !== undefined && token.value !== null)
            ? token.value
            : token.name;

        return v === '%';
    }

    parse(tokens) {
        this.tokens = tokens;
        this.currentTokenIndex = 0;
        this.AST = null;
        this.highestPriorityNode = null;
		this.seenCells.clear();

        if (!tokens || tokens.length == 0)
            return;

        this.AST = this.parseExpression();

        if (this.hasMoreTokens())
            throw new Error("Не все токены были обработаны");

        this.findHighestPriorityOperation(this.AST);
        return this.AST;
    }
	
	revert() {
		this.parse(this.tokens);
	}

    parseExpression(minPrecedence = 0) {
		let left = this.parsePrimary();

		while (this.hasMoreTokens()) {
			const token = this.peekToken();
			if (!this.isPercentToken(token))
				break;

			const opTok = this.consumeToken();
			left = new UnaryOperationNode(opTok, left, this.options);
		}
		
		while (this.hasMoreTokens()) {
			const token = this.peekToken();

			if (token.type !== 9)
				break;

			if (this.isPercentToken(token))
				break;

			const opName = token.value || token.name;
			const operatorInfo = this.operators[opName] || this.operators[token.name];
			if (!operatorInfo || operatorInfo.precedence < minPrecedence)
				break;

			const opTok = this.consumeToken();
			const nextMinPrecedence =
				operatorInfo.associativity === 'right'
					? operatorInfo.precedence
					: operatorInfo.precedence + 1;

			const right = this.parseExpression(nextMinPrecedence);
			left = new BinaryOperationNode(opTok, left, right, this.options);
		}
		
		return left;
	}

    parsePrimary() {
        if (!this.hasMoreTokens())
            throw new Error("Ожидается выражение");

        const token = this.consumeToken();

        if (token.type === -1) return new DefaultNode(token, this.options);
        if (token.type === 0) return new NumberNode(token, this.options);
        if (token.type === 1) return new StringNode(token, this.options);
        if (token.type === 2) return new BooleanNode(token, this.options);
        if (token.type === 3) return new ErrorNode(token, this.options);
        if (token.type === 5 || token.type === 13)
            return new CellRangeNode(token, this.options, this.getColor(token));
        if (token.type === 6 || token.type === 12)
            return new CellNode(token, this.options, this.getColor(token));
        if (token.type === 10 || token.type === 15)
            return new NamedRangeNode(token, this.options, this.getColor(token));
		if (token.type === 11) return new ArrayNode(token, this.options);


        if (token.type === 8) {
            this.expectToken(9, '(');

            const args = [];
            while (this.hasMoreTokens() &&
                this.tokens[this.currentTokenIndex].name !== ')') {
                args.push(this.parseExpression());
            }

            this.expectToken(9, ')');
            return new FunctionCallNode(token, args, this.options);
        }

        if (token.type === 9) {
            const nameOrValue = token.name ?? token.value;

            if (nameOrValue === '(') {
                const expr = this.parseExpression();
                this.expectToken(9, ')');
                return new ParenthesisNode(expr);
            }

            if (nameOrValue === 'un_minus' || nameOrValue === 'un_plus') {
                const operand = this.parsePrimary();
                return new UnaryOperationNode(token, operand, this.options);
            }

            this.currentTokenIndex--;
            return this.parsePrimary();
        }

        throw new Error(`Неожиданный токен: ${token.type} "${token.value}"`);
    }

    hasMoreTokens() {
        return this.currentTokenIndex < this.tokens.length;
    }

    peekToken() {
        return this.hasMoreTokens() ? this.tokens[this.currentTokenIndex] : null;
    }

    consumeToken() {
        if (!this.hasMoreTokens())
            throw new Error("Неожиданный конец выражения");

        return this.tokens[this.currentTokenIndex++];
    }

    expectToken(expectedType, expectedValue = null) {
        const token = this.peekToken();
        const nameOrValue = token ? (token.name ?? token.value) : null;

        if (!token ||
            token.type !== expectedType ||
            (expectedValue !== null && nameOrValue !== expectedValue)) {

            throw new Error(`Ожидается ${expectedType}${
                expectedValue ? ` "${expectedValue}"` : ''
            }, но получен ${token ? token.type : 'конец выражения'}`);
        }

        return this.consumeToken();
    }

	findHighestPriorityOperation(node) {
		this.highestPriorityNode = null;
		if (!node) return null;
		
		let highestPriorityOp = null;
		let maxPrecedence = -1;
		let maxLevel = -1;

		
		this.traverseAST(node, 0, (currentNode, level) => {
			if (currentNode.type === 'BINARY_OPERATION') {
				if (level > maxLevel) {
					maxLevel = level;
					highestPriorityOp = currentNode;
					maxPrecedence = this.operators[currentNode.operator.name].precedence || 0;
					this.highestPriorityNode = highestPriorityOp;
				} else {
					const precedence = this.operators[currentNode.operator.name].precedence || 0;
					if (precedence > maxPrecedence && level == maxLevel) {
						maxPrecedence = precedence;
						highestPriorityOp = currentNode;
						this.highestPriorityNode = highestPriorityOp;
					}
				}
			}
			else if (currentNode.type === 'PARENTHESIS') {
				if (level > maxLevel) {
					maxLevel = level;
					highestPriorityOp = currentNode;
					this.highestPriorityNode = highestPriorityOp;
				}
			}
			else if (currentNode.type === 'FUNCTION_CALL') {
				if (level > maxLevel) {
					maxLevel = level;
					highestPriorityOp = currentNode;
					this.highestPriorityNode = highestPriorityOp;
				}
			}
			else if (currentNode.type === 'UNARY_OPERATION' && currentNode.operator === '%') {
				if (level > maxLevel) {
					maxLevel = level;
					highestPriorityOp = currentNode;
					this.highestPriorityNode = highestPriorityOp;
				}
			}
			else if ((currentNode.type === 'CELL' || currentNode.type === 'NAME')&&
					(!currentNode.parent || (currentNode.parent && currentNode.parent.type != 'FUNCTION_CALL') ||
					(currentNode.parent && (currentNode.parent.type == 'FUNCTION_CALL' && currentNode.parent.arguments.length == 1 && ['ROW','ROWS','COLUMN','FORMULATEXT'].indexOf(currentNode.parent.token.name) == -1)))) {
				if (level > maxLevel) {
					maxLevel = level;
					highestPriorityOp = currentNode;
					this.highestPriorityNode = highestPriorityOp;
				}
			}
		});

		if (!highestPriorityOp && node.type === NodeType.UNARY_OPERATION) {
			highestPriorityOp = node;
		}

		this.highestPriorityNode = highestPriorityOp;
	}

    isInsideFunctionArg(node, funcNode) {
        if (!node || !funcNode) return false;

        let p = node.parent;
        while (p) {
            if (p === funcNode) return true;
            p = p.parent;
        }

        return false;
    }

    traverseAST(node, level, callback) {
        if (!node) return;

        callback(node, level);

        switch (node.type) {
            case 'BINARY_OPERATION':
                this.traverseAST(node.left, level + 1, callback);
                this.traverseAST(node.right, level + 1, callback);
                break;

            case 'UNARY_OPERATION':
                this.traverseAST(node.argument, level + 1, callback);
                break;

            case 'FUNCTION_CALL':
                if (node.arguments) {

                    if (node.name === 'IF' &&
                        node.arguments.length >= 1 &&
                        node.arguments[0].type === NodeType.BOOLEAN) {

                        const cond = node.arguments[0];
                        let branchIndex = -1;

                        if (cond.value === true && node.arguments[1]) {
                            branchIndex = 1;
                        } else if (cond.value === false && node.arguments[2]) {
                            branchIndex = 2;
                        }

                        if (branchIndex !== -1) {
                            const arg = node.arguments[branchIndex];

                            const prev = this.lockFunctionScope;
                            this.lockFunctionScope = node;

                            this.traverseAST(arg, level + 1, callback);

                            this.lockFunctionScope = prev;

                            if (this.highestPriorityNode &&
                                this.isInsideFunctionArg(this.highestPriorityNode, node)) {
                                return;
                            }
                        }
                        return;
                    }

                    for (const arg of node.arguments) {
                        const prev = this.lockFunctionScope;
                        this.lockFunctionScope = node;

                        this.traverseAST(arg, level + 1, callback);

                        this.lockFunctionScope = prev;

                        if (this.highestPriorityNode &&
                            this.isInsideFunctionArg(this.highestPriorityNode, node)) {
                            return;
                        }
                    }
                }
                break;

            case 'PARENTHESIS':
                if (node.expression) {
                    this.traverseAST(node.expression, level + 1, callback);
                }
                break;
        }
    }
	
	getColor(token) {
		let  reference = null;
		if (token.type === 5 || token.type === 6 || token.type === 12 || token.type === 13) {
            reference = (token.ws ? token.ws.sName : token.wsFrom ? token.wsFrom.sName : '' ) + '!' + token.value;
        }
		else if(token.type == 10 || token.type === 15)
			reference = (token.ws ? token.ws.sName : token.wsFrom ? token.wsFrom.sName : '' ) + '!' + token.toRef().value.replaceAll('$',''); 
		if (!reference)
            return null;
		if(!this.seenCells.get(reference)) {
			const color = this.options.formulaRangeBorderColor[this.seenCells.size % this.options.formulaRangeBorderColor.length];
			this.seenCells.set(reference,`rgba(${color.r},${color.g},${color.b},${color.a})`);
		}
		return this.seenCells.get(reference);
	}
	
	toHtml() {
		if (!this.AST) return;

		const output = document.createElement('div');
		output.className = 'output';
		output.setAttribute('id', 'output-' + this.id);
		const address = document.createElement('span');
		address.className = 'address';
		const address_container = document.createElement('span');
		address_container.className = 'container';
		const _sheetName = document.createElement('span');
		_sheetName.className = 'sheet';
		_sheetName.textContent = this.cell.GetWorksheet().GetName();
		if(_sheetName.textContent.indexOf(' ') >= 0) _sheetName.textContent = '\'' + _sheetName.textContent + '\'';
		const divider = document.createElement('span');
		divider.className = 'div';
		divider.textContent = '!';
		const _cellAddress = document.createElement('span');
		_cellAddress.className = 'ref';
		_cellAddress.textContent = this.cell.Address;
		address_container.appendChild(_sheetName);
		address_container.appendChild(divider);
		address_container.appendChild(_cellAddress);
		address.appendChild(address_container);
		output.appendChild(address);
		const equalSign = document.createElement('span');
		equalSign.className = 'equal';
		equalSign.textContent = "=";
		output.appendChild(equalSign);
		const formula_container = document.createElement('span');
		formula_container.className = 'formula';
		const html = this.AST.toHtml();
		formula_container.appendChild(html);
		output.appendChild(formula_container);
		this.el = output;

		output.addEventListener('mouseover', function(e) {
			const span = e.target.closest('span.expression');
			if (span) {
				const hoveredSpans = document.querySelectorAll('span.expression.hover');
				hoveredSpans.forEach(s => s.classList.remove('hover'));
				span.classList.add('hover');
				e.stopPropagation();
			}
		});

		output.addEventListener('mouseout', function(e) {
			const hoveredSpans = document.querySelectorAll('span.expression.hover');
			hoveredSpans.forEach(s => s.classList.remove('hover'));
		});

		if (this.highestPriorityNode && this.highestPriorityNode.el) {
			this.highestPriorityNode.el.classList.add('next');
		}
		return output;
	}
}


class ExpressionProcessor {
    constructor(options = {}) {
		this.parsers = [];
        this.parser = null;
        this.el = null;
		this.options = options;
		this.outputElementId = options.outputElementId;
    }

    process(tokens, cell) {
        try {
            this.parsers.push(new ASTParser(cell, this.options));
			this.parser = this.parsers[this.parsers.length - 1];
			this.parser.parse(tokens);

            if (this.outputElementId && this.parser.AST) {
                this.render();
				var thisIndex = this.parsers.indexOf(this.parser);
				if(thisIndex >= 1) {
					var output = document.getElementById('output-' + this.parsers[thisIndex-1].id);
					var address = output.querySelector('.address .container');
					var link = document.createElement('span');
					link.className = 'link';
					link.style.borderColor = this.parsers[thisIndex-1].highestPriorityNode.color;
					link.style.backgroundColor = this.parsers[thisIndex-1].highestPriorityNode.color;
					var line = document.createElement('span');
					line.className = 'line';
					link.appendChild(line);
					address.appendChild(link);
				}
			}

            return {
                success: true,
                ast: this.parser.AST
            };
        }
		catch (error) {
            console.error('Ошибка обработки выражения:', error);
            return {
                success: false,
                error: error.message
            };
        }
    }

    canCalculate() {
        return this.parser && this.parser.highestPriorityNode;
    }
	
	canRevert() {
		return this.parser && this.parser.AST && this.parser.AST.token != this.parser.tokens[0];
	}
	
	canStepIn() {
		return this.parser && this.parser.highestPriorityNode && ['CELL','NAME'].indexOf(this.parser.highestPriorityNode.type) >= 0 ;
	}
	
	canStepOut() {
		return this.parsers.indexOf(this.parser) > 0;
	}
	
	stepOut() {
		if(this.parser) {
			document.getElementById('output-' + this.parser.id).remove();
			this.parsers.pop();
			this.parser = this.parsers[this.parsers.length - 1];
			var link = this.parser.el.querySelector('.link');
			if(link)
				link.remove();
			this.calculateNextStep();
		}
	}

    calculateNextStep() {
        if (this.parser && this.parser.highestPriorityNode && this.parser.highestPriorityNode.calculate) {
            
			this.parser.el.querySelectorAll('.last').forEach(el => el.classList.remove('last'));
            
			var lastNode = this.parser.highestPriorityNode.calculate();
            
			if (lastNode)
                this.parser.AST = lastNode;
            
			this.parser.findHighestPriorityOperation(this.parser.AST);
            
			if (this.parser.highestPriorityNode && this.parser.highestPriorityNode.el) {
                if (document.getElementById('node-' + this.parser.highestPriorityNode.id)) {
                    document.getElementById('node-' + this.parser.highestPriorityNode.id).classList.add('next');
                }
            }
        }
    }
	
	revert() {
		if(this.parser) {
			this.parser.revert();
			document.getElementById('output-' + this.parser.id).remove();
			this.render();
		}
	}

    render() {
		if(this.parser) {
			const container = document.getElementById(this.outputElementId);
			if (!container) {
				console.error(`Элемент с id "${this.outputElementId}" не найден`);
				return;
			}
			var output = this.parser.toHtml();
			container.appendChild(output);
			output.scrollIntoView({behavior: "smooth", block: "end", inline: "nearest"});
			this.parser.el = document.getElementById('output-' + this.parser.id);
		}
    }
	
	clear() {
		this.parsers.length = 0;
		this.parser = null;
		const container = document.getElementById(this.outputElementId);
		if (container)
			container.innerHTML = ''
	}
}
