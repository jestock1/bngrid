function BNGrid() {
	var TOK_TYPE_NOOP			= "noop";
	var TOK_TYPE_OPERAND		= "operand";
	var TOK_TYPE_FUNCTION		= "function";
	var TOK_TYPE_SUBEXPR		= "subexpression";
	var TOK_TYPE_ARGUMENT		= "argument";
	var TOK_TYPE_OP_PRE			= "operator-prefix";
	var TOK_TYPE_OP_IN			= "operator-infix";
	var TOK_TYPE_OP_POST		= "operator-postfix";
	var TOK_TYPE_WSPACE			= "white-space";
	var TOK_TYPE_UNKNOWN		= "unknown"

	var TOK_SUBTYPE_START		= "start";
	var TOK_SUBTYPE_STOP		= "stop";

	var TOK_SUBTYPE_TEXT		= "text";
	var TOK_SUBTYPE_NUMBER		= "number";
	var TOK_SUBTYPE_LOGICAL		= "logical";
	var TOK_SUBTYPE_ERROR		= "error";
	var TOK_SUBTYPE_RANGE		= "range";

	var TOK_SUBTYPE_MATH		= "math";
	var TOK_SUBTYPE_CONCAT		= "concatenate";
	var TOK_SUBTYPE_INTERSECT	= "intersect";
	var TOK_SUBTYPE_UNION		= "union";

	this.sheetIndexCounter = 0;
	this.rows = new Array();
	this.rows.push (
		new Array (
			"Sheet " + (++this.sheetIndexCounter),	// Worksheet name
			new Array(),							// Columns
			new Array()								// Forumlas for each column
		)
	);
	this.attributes = new Array();
	this.selectedColumn = -1;
	this.selectedRow = -1;
	this.selectedSheet = 0;
	this.selectedCell = undefined;
	this.editing = false;

	function parseFormula(formula) {
		var indentCount = 0;
		var ret = new Array();

		var indent = function() {
			var s = "|";
			for (var i = 0; i < indentCount; i++) {
				s += "&nbsp;&nbsp;&nbsp;|";
			}
			return s;
		};

		var tokens = getTokens(formula);
		var tokensHtml = "";
		tokensHtml += "<table cellspacing=\"0\" style=\"border-top: 1px #cecece solid; margin-top: 5px; margin-bottom: 5px\">";
		tokensHtml += "<tr>";
		tokensHtml += "<td class=\"token\" style=\"font-weight: bold; width: 50px\">index</td>";
		tokensHtml += "<td class=\"token\" style=\"font-weight: bold; width: 125px\">type</td>";
		tokensHtml += "<td class=\"token\" style=\"font-weight: bold; width: 125px\">subtype</td>";
		tokensHtml += "<td class=\"token\" style=\"font-weight: bold; width: 150px\">token</td>";
		tokensHtml += "<td class=\"token\" style=\"font-weight: bold; width: 300px\">token tree</td></tr>";

		while (tokens.moveNext()) {
			var token = tokens.current();

			if (token.subtype == TOK_SUBTYPE_STOP)
				indentCount -= ((indentCount > 0) ? 1 : 0);

			tokensHtml += "<tr>";
			tokensHtml += "<td class=\"token\">" + (tokens.index + 1) + "</td>";
			tokensHtml += "<td class=\"token\">" + token.type + "</td>";
			tokensHtml += "<td class=\"token\">" + ((token.subtype.length == 0) ? "&nbsp;" : token.subtype) + "</td>";
			tokensHtml += "<td class=\"token\">" + ((token.value.length == 0) ? "&nbsp;" : token.value).split(" ").join("&nbsp;") + "</td>";
			tokensHtml += "<td class=\"token\">" + indent() + ((token.value.length == 0) ? "&nbsp;" : token.value).split(" ").join("&nbsp;") + "</td>";
			tokensHtml += "</tr>";

			ret.push (
				new Array (
					tokens.index + 1,
					token.type,
					((token.subtype.length == 0) ? "" : token.subtype),
					(token.value.length == 0) ? "" : token.value
				)
			);

			if (token.subtype == TOK_SUBTYPE_START)
				indentCount += 1;
		}

		tokensHtml += "</table>";

		return ret;
	}


	function f_token(value, type, subtype) {
		this.value = value;
		this.type = type;
		this.subtype = subtype;
	}

	function f_tokens() {
		this.items = new Array();

		this.add = function(value, type, subtype) {
			if (!subtype)
				subtype = "";
			token = new f_token(value, type, subtype);
			this.addRef(token);
			return token;
		};

		this.addRef = function(token) {
			this.items.push(token);
		};

		this.index = -1;
		this.reset = function() { this.index = -1; };
		this.BOF = function() { return (this.index <= 0); };
		this.EOF = function() { return (this.index >= (this.items.length - 1)); };
		this.moveNext = function() { if (this.EOF()) return false; this.index++; return true; };
		this.current = function() { if (this.index == -1) return null; return (this.items[this.index]); };
		this.next = function() { if (this.EOF()) return null; return (this.items[this.index + 1]); };
		this.previous = function() { if (this.index < 1) return null; return (this.items[this.index - 1]); };
	}

	function f_tokenStack() {
		this.items = new Array();
	
		this.push = function(token) { this.items.push(token); };
		this.pop = function() { var token = this.items.pop(); return (new f_token("", token.type, TOK_SUBTYPE_STOP)); };

		this.token = function() { return ((this.items.length > 0) ? this.items[this.items.length - 1] : null); };
		this.value = function() { return ((this.token()) ? this.token().value : ""); };
		this.type = function() { return ((this.token()) ? this.token().type : ""); };
		this.subtype = function() { return ((this.token()) ? this.token().subtype : ""); };
	}

	function getTokens(formula) {
		var tokens = new f_tokens();
		var tokenStack = new f_tokenStack();
		var offset = 0;
		var currentChar = function() { return formula.substr(offset, 1); };
		var doubleChar  = function() { return formula.substr(offset, 2); };
		var nextChar    = function() { return formula.substr(offset + 1, 1); };
		var EOF         = function() { return (offset >= formula.length); };
		var token = "";
		var inString = false;
		var inPath = false;
		var inRange = false;
		var inError = false;

		while (formula.length > 0) {
			if (formula.substr(0, 1) == " ")
				formula = formula.substr(1);
			else {
				if (formula.substr(0, 1) == "=")
					formula = formula.substr(1);
				break;
			}
		}

		var regexSN = /^[1-9]{1}(\.[0-9]+)?E{1}$/;
		while (!EOF()) {
			// state-dependent character evaluation (order is important)

			// double-quoted strings
			// embeds are doubled
			// end marks token
			if (inString) {
				if (currentChar() == "\"") {
					if (nextChar() == "\"") {
						token += "\"";
						offset += 1;
					}
					else {
						inString = false;
						tokens.add(token, TOK_TYPE_OPERAND, TOK_SUBTYPE_TEXT);
						token = "";
					}
				}
				else {
					token += currentChar();
				}
				offset += 1;
				continue;
			}

			// single-quoted strings (links)
			// embeds are double
			// end does not mark a token
			if (inPath) {
				if (currentChar() == "'") {
					if (nextChar() == "'") {
						token += "'";
						offset += 1;
					}
					else {
						inPath = false;
					}
				}
				else {
					token += currentChar();
				}
				offset += 1;
				continue;
			}

			// bracked strings (range offset or linked workbook name)
			// no embeds (changed to "()" by Excel)
			// end does not mark a token

			if (inRange) {
				if (currentChar() == "]") {
					inRange = false;
				}
				token += currentChar();
				offset += 1;
				continue;
			}

			// error values
			// end marks a token, determined from absolute list of values
			if (inError) {
				token += currentChar();
				offset += 1;
				if ((",#NULL!,#DIV/0!,#VALUE!,#REF!,#NAME?,#NUM!,#N/A,").indexOf("," + token + ",") != -1) {
					inError = false;
					tokens.add(token, TOK_TYPE_OPERAND, TOK_SUBTYPE_ERROR);
					token = "";
				}
				continue;
			}

			// scientific notation check
			if (("+-").indexOf(currentChar()) != -1) {
				if (token.length > 1) {
					if (token.match(regexSN)) {
						token += currentChar();
						offset += 1;
						continue;
					}
				}
			}

			// independent character evaulation (order not important)

			// establish state-dependent character evaluations
			if (currentChar() == "\"") {
				if (token.length > 0) {
					// not expected
					tokens.add(token, TOK_TYPE_UNKNOWN);
					token = "";
				}
				inString = true;
				offset += 1;
				continue;
			}

			if (currentChar() == "'") {
				if (token.length > 0) {
					// not expected
					tokens.add(token, TOK_TYPE_UNKNOWN);
					token = "";
				}
				inPath = true;
				offset += 1;
				continue;
			}

			if (currentChar() == "[") {
				inRange = true;
				token += currentChar();
				offset += 1;
				continue;
			}

			if (currentChar() == "#") {
				if (token.length > 0) {
					// not expected
					tokens.add(token, TOK_TYPE_UNKNOWN);
					token = "";
				}
				inError = true;
				token += currentChar();
				offset += 1;
				continue;
			}

			// mark start and end of arrays and array rows
			if (currentChar() == "{") {
				if (token.length > 0) {
					// not expected
					tokens.add(token, TOK_TYPE_UNKNOWN);
					token = "";
				}
				tokenStack.push(tokens.add("ARRAY", TOK_TYPE_FUNCTION, TOK_SUBTYPE_START));
				tokenStack.push(tokens.add("ARRAYROW", TOK_TYPE_FUNCTION, TOK_SUBTYPE_START));
				offset += 1;
				continue;
			}

			if (currentChar() == ";") {
				if (token.length > 0) {
					tokens.add(token, TOK_TYPE_OPERAND);
					token = "";
				}
				tokens.addRef(tokenStack.pop());
				tokens.add(",", TOK_TYPE_ARGUMENT);
				tokenStack.push(tokens.add("ARRAYROW", TOK_TYPE_FUNCTION, TOK_SUBTYPE_START));
				offset += 1;
				continue;
			}

			if (currentChar() == "}") {
				if (token.length > 0) {
					tokens.add(token, TOK_TYPE_OPERAND);
					token = "";
				}
				tokens.addRef(tokenStack.pop());
				tokens.addRef(tokenStack.pop());
				offset += 1;
				continue;
			}

			// trim white-space
			if (currentChar() == " ") {
				if (token.length > 0) {
					tokens.add(token, TOK_TYPE_OPERAND);
					token = "";
				}
				tokens.add("", TOK_TYPE_WSPACE);
				offset += 1;
				while ((currentChar() == " ") && (!EOF())) {
					offset += 1;
				}
				continue;
			}

			// multi-character comparators
			if ((",>=,<=,<>,").indexOf("," + doubleChar() + ",") != -1) {
				if (token.length > 0) {
					tokens.add(token, TOK_TYPE_OPERAND);
					token = "";
				}
				tokens.add(doubleChar(), TOK_TYPE_OP_IN, TOK_SUBTYPE_LOGICAL);
				offset += 2;
				continue;
			}

			// standard infix operators
			if (("+-*/^&=><").indexOf(currentChar()) != -1) {
				if (token.length > 0) {
					tokens.add(token, TOK_TYPE_OPERAND);
					token = "";
				}
				tokens.add(currentChar(), TOK_TYPE_OP_IN);
				offset += 1;
				continue;
			}

			// standard postfix operators
			if (("%").indexOf(currentChar()) != -1) {
				if (token.length > 0) {
					tokens.add(token, TOK_TYPE_OPERAND);
					token = "";
				}
				tokens.add(currentChar(), TOK_TYPE_OP_POST);
				offset += 1;
				continue;
			}

			// start subexpression or function
			if (currentChar() == "(") {
				if (token.length > 0) {
					tokenStack.push(tokens.add(token, TOK_TYPE_FUNCTION, TOK_SUBTYPE_START));
					token = "";
				}
				else {
					tokenStack.push(tokens.add("", TOK_TYPE_SUBEXPR, TOK_SUBTYPE_START));
				}
				offset += 1;
				continue;
			}

			// function, subexpression, array parameters
			if (currentChar() == ",") {
				if (token.length > 0) {
					tokens.add(token, TOK_TYPE_OPERAND);
					token = "";
				}
				if (!(tokenStack.type() == TOK_TYPE_FUNCTION)) {
					tokens.add(currentChar(), TOK_TYPE_OP_IN, TOK_SUBTYPE_UNION);
				}
				else {
					tokens.add(currentChar(), TOK_TYPE_ARGUMENT);
				}
				offset += 1;
				continue;
			}

			// stop subexpression
			if (currentChar() == ")") {
				if (token.length > 0) {
					tokens.add(token, TOK_TYPE_OPERAND);
					token = "";
				}
				tokens.addRef(tokenStack.pop());
				offset += 1;
				continue;
			}

			// token accumulation
			token += currentChar();
			offset += 1;
		}

		// dump remaining accumulation
		if (token.length > 0)
			tokens.add(token, TOK_TYPE_OPERAND);

		// move all tokens to a new collection, excluding all unnecessary white-space tokens
		var tokens2 = new f_tokens();

		while (tokens.moveNext()) {
			token = tokens.current();

			if (token.type == TOK_TYPE_WSPACE) {
				if ((tokens.BOF()) || (tokens.EOF())) {
					;
				}
				else if (!(((tokens.previous().type == TOK_TYPE_FUNCTION) && (tokens.previous().subtype == TOK_SUBTYPE_STOP))
						|| ((tokens.previous().type == TOK_TYPE_SUBEXPR) && (tokens.previous().subtype == TOK_SUBTYPE_STOP))
						|| (tokens.previous().type == TOK_TYPE_OPERAND))) {
					;
				}
				else if (!(((tokens.next().type == TOK_TYPE_FUNCTION) && (tokens.next().subtype == TOK_SUBTYPE_START))
						|| ((tokens.next().type == TOK_TYPE_SUBEXPR) && (tokens.next().subtype == TOK_SUBTYPE_START))
						|| (tokens.next().type == TOK_TYPE_OPERAND))) {
					;
				}
				else
					tokens2.add(token.value, TOK_TYPE_OP_IN, TOK_SUBTYPE_INTERSECT);
				continue;
			}

			tokens2.addRef(token);
		}

		// switch infix "-" operator to prefix when appropriate, switch infix "+" operator to noop when appropriate, identify operand
		// and infix-operator subtypes, pull "@" from in front of function names
		while (tokens2.moveNext()) {
			token = tokens2.current();
			if ((token.type == TOK_TYPE_OP_IN) && (token.value == "-")) {
				if (tokens2.BOF())
					token.type = TOK_TYPE_OP_PRE;
				else if (((tokens2.previous().type == TOK_TYPE_FUNCTION) && (tokens2.previous().subtype == TOK_SUBTYPE_STOP))
						|| ((tokens2.previous().type == TOK_TYPE_SUBEXPR) && (tokens2.previous().subtype == TOK_SUBTYPE_STOP))
						|| (tokens2.previous().type == TOK_TYPE_OP_POST)
						|| (tokens2.previous().type == TOK_TYPE_OPERAND))
					token.subtype = TOK_SUBTYPE_MATH;
				else
					token.type = TOK_TYPE_OP_PRE;
				continue;
			}

			if ((token.type == TOK_TYPE_OP_IN) && (token.value == "+")) {
				if (tokens2.BOF())
					token.type = TOK_TYPE_NOOP;
				else if (((tokens2.previous().type == TOK_TYPE_FUNCTION) && (tokens2.previous().subtype == TOK_SUBTYPE_STOP))
						|| ((tokens2.previous().type == TOK_TYPE_SUBEXPR) && (tokens2.previous().subtype == TOK_SUBTYPE_STOP))
						|| (tokens2.previous().type == TOK_TYPE_OP_POST)
						|| (tokens2.previous().type == TOK_TYPE_OPERAND))
					token.subtype = TOK_SUBTYPE_MATH;
				else
					token.type = TOK_TYPE_NOOP;
				continue;
			}

			if ((token.type == TOK_TYPE_OP_IN) && (token.subtype.length == 0)) {
				if (("<>=").indexOf(token.value.substr(0, 1)) != -1)
					token.subtype = TOK_SUBTYPE_LOGICAL;
				else if (token.value == "&")
					token.subtype = TOK_SUBTYPE_CONCAT;
				else
					token.subtype = TOK_SUBTYPE_MATH;
				continue;
			}

			if ((token.type == TOK_TYPE_OPERAND) && (token.subtype.length == 0)) {
				if (isNaN(parseFloat(token.value))) {
					if ((token.value == "TRUE") || (token.value == "FALSE"))
						token.subtype = TOK_SUBTYPE_LOGICAL;
					else
						token.subtype = TOK_SUBTYPE_RANGE;
				}
				else {
					token.subtype = TOK_SUBTYPE_NUMBER;
				}
				continue;
			}

			if (token.type == TOK_TYPE_FUNCTION) {
				if (token.value.substr(0, 1) == "@")
					token.value = token.value.substr(1);
				continue;
			}
		}

		tokens2.reset();

		// move all tokens to a new collection, excluding all noops
		tokens = new f_tokens();

		while (tokens2.moveNext()) {
			if (tokens2.current().type != TOK_TYPE_NOOP)
				tokens.addRef(tokens2.current());
		}

		tokens.reset();
		return tokens;
	}

	this.addSheet = function () {
		this.rows.push (
			new Array (
				"Sheet " + (++this.sheetIndexCounter),
				new Array(),
				new Array()
			)
		);
		this.draw();
	}

	this.delSheet = function (sheetIndex) {
		if (this.rows.length == 1) {
			alert("You cannot delete the last worksheet");
			return;
		}

		var rows = new Array();

		sheetIndex = parseInt(sheetIndex);
		var i;
		for (i = 0; i < this.rows.length; i++) {
			if (i != sheetIndex)
				rows.push(this.rows[i]);
		}

		this.rows = rows;
		this.draw();
	}

	this.selectSheet = function (index) {
		this.selectedSheet = parseInt(index);
		this.draw();
	}

	this.add = function () {
		if (this.rows.length === undefined) {
			this.rows.push (
				new Array (
					"Sheet " + (++this.sheetIndexCounter),
					new Array(),
					new Array()
				)
			);
		}

		var rows = this.rows[this.selectedSheet][1];
		var formulas = this.rows[this.selectedSheet][2];

		rows.push(arguments);

		var i;
		var ar = new Array();
		for (i = 0; i < arguments.length; i++)
			ar.push("");

		formulas.push(ar);
	}

	this.getRows = function (sheetIndex) {
		try {
			return this.rows[sheetIndex][1];
		}
		catch (ex) {
			;
		}
	}

	this.getFormulas = function (sheetIndex) {
		return this.rows[sheetIndex][2];
	}

	this.getFormula = function (row, col) {
		var formulas = this.getFormulas(this.selectedSheet);

		if ((formulas[row] !== undefined) && (formulas[row][col] !== undefined) && (formulas[row][col].length > 0))
			return formulas[row][col];

		return "";
	}

	this.setFormula = function (row, col, value) {
		var formulas = this.getFormulas(this.selectedSheet);

		if ((formulas[row] !== undefined) && (formulas[row][col] !== undefined))
			formulas[row][col] = value;
	}

	this.setValue = function (row, col, value) {
		var rows = this.getRows(this.selectedSheet);

		if ((rows[row] !== undefined) && (rows[row][col] !== undefined))
			rows[row][col] = value;
	}

	this.getValue = function (row, col) {
		var rows = this.getRows(this.selectedSheet);

		if ((rows[row] !== undefined) && (rows[row][col] !== undefined))
			return rows[row][col];
	}

	this.highlightColumn = function (column) {
		var i;
		var rows = this.getRows(this.selectedSheet);
		for (i = 0; i < rows.length; i++) {
			var el = document.getElementById("rowh" + (i + 1));
			el.setAttribute("style", "background-color: #e2ebf2");

			var j;
			for (j = 0; j < rows[i].length; j++) {
				var el = document.getElementById("colh" + (j + 1));
				el.setAttribute("style", "background-color: #e2ebf2");
			}
		}

		for (i = 0; i < rows.length; i++) {
			var children = document.getElementById("row" + (i + 1)).childNodes;

			var j;
			for (j = 0; j < children.length; j++) {
				var id = children[j].getAttribute("id");
				if (id !== undefined && id !== null && j == parseInt(column) && id != ("rowh" + (i + 1))) {
					var border = "";
					children[j].setAttribute("style", "background-color: #fbc271; " + border);
				}
				else if (id !== undefined && id !== null && id != "" && id != ("rowh" + (i + 1))) {
					var border = "";
					children[j].setAttribute("style", "background-color: #FFFFFF; " + border);
				}
			}
		}

		var id = document.getElementById("colh" + column);
		id.setAttribute("style", "background: url(); background-color: #fbc271;");

		this.selectedRow = -1;
		this.selectedColumn = parseInt(column);
	}

	this.highlightRow = function (row) {
		row = row.replace(/^row/, "");

		var children = document.getElementById("row" + row).childNodes;

		var i;
		var rows = this.getRows(this.selectedSheet);
		for (i = 0; i < rows.length; i++) {
			var children2 = document.getElementById("row" + (i + 1)).childNodes;
			var j;
			for (j = 1; j < children2.length; j++)
				children2[j].setAttribute("style", "background-color: #FFFFFF;");
		}

		for (i = 1; i < children.length; i++)
			children[i].setAttribute("style", "background-color: #fbc271;");

		for (i = 0; i < rows.length; i++) {
			if ((i + 1) != parseInt(row)) {
				var el = document.getElementById("rowh" + (i + 1));
				el.setAttribute("style", "background-color: #e2ebf2");

				var j;
				for (j = 0; j < rows[i].length; j++) {
					var el = document.getElementById("colh" + (j + 1));
					el.setAttribute("style", "background-color: #e2ebf2");
				}
			}
		}

		var id = document.getElementById("rowh" + row);
		id.setAttribute("style", "background-color: #fbc271;");

		this.selectedColumn = -1;
		this.selectedRow = parseInt(row);
	}

	this.setCell = function (cell, value) {
		if ((cell === null) || (cell === undefined))
			return;

		var id = cell.getAttribute("id");

		var arr = id.split(";");
		var row = arr[0];
		var col = arr[1];

		row = row.replace(/^row/, "");
		col = col.replace(/^col/, "");

		row = parseInt(row) - 1;
		col = parseInt(col) - 1;

		// If the new value starts with = then we need to parse the formula
		if (value.match(/^=/) && (value.length > 1)) {
			this.setFormula(row, col, value);
			value = this.parseExcelFormula(value);
		}
		else {
			this.setFormula(row, col, "");
		}

		if (document.getElementById) {
			try {
				cell.replaceChild(document.createTextNode(value), cell.firstChild);
			} catch (ex) {
				;
			}

			if (value == "")
				cell.innerHTML += "&nbsp;";
		}
		else if (document.all) {
			cell.innerText = value;

			if (value == "")
				cell.innerHTML += "&nbsp;";
		}

		var id = cell.getAttribute("id");

		var arr = id.split(";");
		var row = arr[0];
		var col = arr[1];

		row = row.replace(/^row/, "");
		col = col.replace(/^col/, "");

		row = parseInt(row) - 1;
		col = parseInt(col) - 1;

		var rows = this.rows[this.selectedSheet][1];

		rows[row][col] = value;

		if (cell.setAttribute) {
			cell.setAttribute (
				"style",
				"border: 1px solid #000000; " +
				"border-top: 2px solid #000000; " +
				"border-left: 2px solid #000000;"
			);
		}

		this.draw();
		this.selectedCell = id;
		this.highlightHeader(id);
		this.highlightCell(document.getElementById(id), this);
		var formula = this.getFormula(row, col);

		if (formula.length == 0)
			formula = this.getValue(row, col);
		document.getElementById("formulaEditBox").setAttribute("value", formula);
		document.getElementById("formulaEditBox").focus();
	}

	this.translateCellReference = function (ref) {
		var rows = this.getRows(this.selectedSheet);

		ref = ref.toUpperCase();

		var i;
		if ((rows.length == 0) || (rows[0].length == 0)) {
			alert("Invalid cell reference in formula.");
			return undefined;
		}

		var cols = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

		var colRef = new Array();
		for (i = 0; i < ref.length; i++) {
			if (cols.search(ref.charAt(i)) != -1)
				colRef.push(ref.charAt(i));
		}

		var col = cols.search(colRef[0]);
		if (colRef.length > 1)
			col = ((colRef.length - 1) * cols.length) + cols.search(colRef.length - 1);

		var row = parseInt(ref.substr(colRef.length)) - 1;

		if ((rows[row] === undefined) || (rows[row][col] === undefined))
			return undefined;

		return rows[row][col];
	}

	this.getRowReferenceIndex = function (ref) {
		var rows = this.getRows(this.selectedSheet);

		ref = ref.toUpperCase();

		var i;
		if ((rows.length == 0) || (rows[0].length == 0)) {
			alert("Invalid cell reference in formula.");
			return undefined;
		}

		var cols = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

		var colRef = new Array();
		for (i = 0; i < ref.length; i++) {
			if (cols.search(ref.charAt(i)) != -1)
				colRef.push(ref.charAt(i));
		}

		var col = cols.search(colRef[0]);
		if (colRef.length > 1)
			col = ((colRef.length - 1) * cols.length) + cols.search(colRef.length - 1);

		var row = parseInt(ref.substr(colRef.length)) - 1;

		return row;
	}

	this.getCellReferenceIndex = function (ref) {
		var rows = this.getRows(this.selectedSheet);

		ref = ref.toUpperCase();

		var i;
		if ((rows.length == 0) || (rows[0].length == 0)) {
			alert("Invalid cell reference in formula.");
			return undefined;
		}

		var cols = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
		var colRef = new Array();
		for (i = 0; i < ref.length; i++) {
			if (cols.search(ref.charAt(i)) != -1)
				colRef.push(ref.charAt(i));
		}

		var col = cols.search(colRef[0]);
		if (colRef.length > 1)
			col = ((colRef.length - 1) * cols.length) + cols.search(colRef.length - 1);

		return col;
	}

	this.getCellReferenceVal = function (ref) {
		var rows = this.getRows(this.selectedSheet);

		ref = ref.toUpperCase();

		var i;
		if ((rows.length == 0) || (rows[0].length == 0)) {
			alert("Invalid cell reference in formula.");
			return undefined;
		}

		var cols = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
		var colRef = "";
		for (i = 0; i < ref.length; i++) {
			if (cols.search(ref.charAt(i)) != -1)
				colRef += ref.charAt(i);
		}

		return colRef;
	}

	this.getInnerSubExpression = function (parseTree) {
    	var i;
    	var ptOps = new Array();
    	for (i = 0; i < parseTree.length; i++) {
      		if (parseTree[i][1] == "subexpression" && parseTree[i][2] == "start" && ((i + 1) < parseTree.length && parseTree[i + 1][2] !== "start")) {
        		var j;

        		ptOps.push(parseTree[i]);
        		i++;
        		for (j = i; j < parseTree.length && parseTree[j][2] !== "stop"; j++, i++)
          			ptOps.push(parseTree[j]);

        		ptOps.push(parseTree[i]);
        		return ptOps;
			}
		}

		// No sub-expressions found. See if there are any math formulas
		// to calculate. I didn't want to iterate through the parse tree
		// again but this method works so I'll leave it for now.
		if (ptOps.length == 0) {
			for (i = 0; i < parseTree.length; i++) {
				// Embedded sub-expression (probably within a function call)
				if (i > 0 && i < (parseTree.length - 1) && parseTree[i][1] == 'operand' && parseTree[i + 1][2] == 'math') {
					// operand
					ptOps.push(parseTree[i]);
					i++;

					var j;
					for (j = i; j < parseTree.length && (parseTree[j][1] == 'operand' || parseTree[j][2] == 'math'); j++)
						ptOps.push(parseTree[j]);
				}
			}
		}

		return ptOps;
	}

	this.getFunction = function (parseTree) {
		var i;
		var start = 0;
		var args = new Array();

		for (i = parseTree.length - 1; i > 0; i--) {
			if ((parseTree[i][1] == "function") && (parseTree[i][2] == "start")) {
				args.push(parseTree[i]);
				var j;
				i++;
				for (j = i; j < parseTree.length && parseTree[j][2] !== "stop"; j++)
					args.push(parseTree[j]);

				args.push(parseTree[j]);
				return args;
			}
		}

		return undefined;
	}

	this.supportedExcelFunctions = new Array (
		// Math Functions
		new Array (
			"ABS",
			"Returns the absolute value of a number",
			"(number)",
			function (exprTree) {
				if (exprTree.length != 1 || parseFloat(exprTree[0][3]).toString() === "NaN") {
					return "#FUNCTION!";
				}
				return Math.abs(parseFloat(exprTree[0][3]));
			}
		),
		new Array (
			"ACOS",
			"Returns the arccosine of a number",
			"(number)",
			function (exprTree) {
				if ((exprTree.length != 1) || (parseFloat(exprTree[0][3]).toString() === "NaN"))
					return "#FUNCTION!";

				var arg = parseFloat(exprTree[0][3]);
				if ((arg < -1) || (arg > 1))
					return "#VAL!";

				return Math.acos(arg);
			}
		),
		new Array (
			"ACOSH",
			"Returns the inverse hyperbolic cosine of a number",
			"(number)",
			function (exprTree) {
				if ((exprTree.length != 1) || (parseInt(exprTree[0][3]).toString() === "NaN"))
					return "#FUNCTION!";

				var arg = parseInt(exprTree[0][3]);

				if (arg < 1)
					return "#VAL!";

				return Math.log(arg + Math.sqrt(arg * arg - 1));
			}
		),
    new Array (
      "ASIN",
      "Returns the arcsine of a number",
      "(number)",
      function (exprTree) {
        if (exprTree.length != 1 || parseInt(exprTree[0][3]).toString() === "NaN") {
          return "#FUNCTION!";
        }

        var arg = parseInt(exprTree[0][3]);

        if (arg < -1 || arg > 1) {
          return "#VAL!";
        }

        return Math.asin(arg);
      }
    ),
    new Array (
      "ASINH",
      "Returns the inverse hyperbolic sine of a number",
      "(number)",
      function (exprTree) {
        if (exprTree.length != 1 || parseInt(exprTree[0][3]).toString() === "NaN") {
          return "#FUNCTION!";
        }

        var arg = parseInt(exprTree[0][3]);

        return Math.log(arg + Math.sqrt(arg * arg + 1));
      }
    ),
    new Array (
      "ATAN",
      "Returns the arctangent of a number",
      "(number)",
      function (exprTree) {
        if (exprTree.length != 1 || parseFloat(exprTree[0][3]).toString() === "NaN") {
          return "#FUNCTION!";
        }

        var arg = parseFloat(exprTree[0][3]);

        return Math.atan(arg);
      }
    ),
    new Array (
      "ATAN2",
      "Returns the arctangent from x- and y-coordinates",
      "(number)",
      function (exprTree) {
        if (exprTree.length != 2 || parseInt(exprTree[0][3]).toString() === "NaN" || parseInt(exprTree[1][3]).toString() === "NaN") {
          return "#FUNCTION!";
        }

        var x = parseInt(exprTree[0][3]);
        var y = parseInt(exprTree[1][3]);

        if (x < -1 || x > 1 || y < -1 || y > 1) {
          return "#VAL!";
        }

        if (x == 0 && y == 0) {
          return "#DIV/0!";
        }

        return Math.atan2(y, x);
      }
    ),
    new Array (
      "ATANH",
      "Returns the inverse hyperbolic tangent of a number",
      "(number)",
      function (exprTree) {
        if (exprTree.length != 1 || parseInt(exprTree[0][3]).toString() === "NaN") {
          return "#FUNCTION!";
        }

        var arg = parseInt(exprTree[0][3]);

        if (arg < -1 || arg > 1) {
          return "#VAL!";
        }

        return 0.5 * Math.log((1 + arg) / (1 - arg));
      }
    ),
    new Array (
      "CEILING",
      "Rounds a number to the nearest integer or to the nearest multiple of significance",
      "(number)",
      function (exprTree) {
        if (exprTree.length != 1 || parseFloat(exprTree[0][3]).toString() === "NaN") {
          return "#FUNCTION!";
        }

        var arg = parseFloat(exprTree[0][3]);

        return Math.ceil(arg);
      }
    ),
    new Array (
      "COMBIN",
      "Returns the number of combinations for a given number of objects",
      function () { }
    ),
    new Array (
      "COS",
      "Returns the cosine of a number",
      "(number)",
      function (exprTree) {
        if (exprTree.length != 1 || parseFloat(exprTree[0][3]).toString() === "NaN") {
          return "#FUNCTION!";
        }

        var arg = parseFloat(exprTree[0][3]);

        return Math.cos(arg);
      }
    ),
    new Array (
      "COSH",
      "Returns the hyperbolic cosine of a number",
      "(number)",
      function (exprTree) {
        if (exprTree.length != 1 || parseInt(exprTree[0][3]).toString() === "NaN") {
          return "#FUNCTION!";
        }

        var arg = parseInt(exprTree[0][3]);

        return (Math.exp(arg) + Math.exp(-arg)) / 2;
      }
    ),
    new Array (
      "DEGREES",
      "Converts radians to degrees",
      "(angle)",
      function (exprTree) {
        if (exprTree.length != 1 || parseInt(exprTree[0][3]).toString() === "NaN") {
          return "#FUNCTION!";
        }

        var arg = parseInt(exprTree[0][3]);

        return (arg / parseFloat(Math.PI.toString().substr(0, 16))) * 180;
      }
    ),
    new Array (
      "EVEN",
      "Rounds a number up to the nearest even integer",
      "(number)",
      function (exprTree) {
        if (exprTree.length != 1) {
          return "#FUNCTION!";
        }

        if (parseInt(exprTree[0][3]).toString() === "NaN") {
          return "#VAL!";
        }

        var neg = false;

        if (parseInt(exprTree[0][3]) < 0) {
          neg = true;
        }

        var ret = Math.round(Math.abs(parseInt(exprTree[0][3])));

        if (ret % 2 != 0) {
          ret++;
        }

        if (neg == true) {
          ret = 0 - ret;
        }

        return ret;
      }
    ),
    new Array (
      "EXP",
      "Returns e raised to the power of a given number",
      function () { }
    ),
    new Array (
      "FACT",
      "Returns the factorial of a number",
      "(number)",
      function (exprTree) {
        if (exprTree.length !=  parseInt(exprTree[0][3]).toString() === "NaN") {
          return "#FUNCTION!";
        }

        var arg = parseInt(exprTree[0][3]);

        if (arg < 0) {
          return "#NUM!";
        }

        var i;
        var result = 1;
        for (i = 2; i <= arg; i++) {
          result *= i
        }
        return result;
      }
    ),
    new Array (
      "FACTDOUBLE",
      "Returns the double factorial of a number",
      function () { }
    ),
    new Array (
      "FLOOR",
      "Rounds a number down, toward zero",
      function (exprTree) { }
    ),
    new Array (
      "GCD",
      "Returns the greatest common divisor",
      "(number1, number2, ...)",
      function (exprTree) {
        if (exprTree.length < 2) {
          return "#FUNCTION!";
        }

        // Parameter sanity
        var i;
        for (i = 0; i < exprTree.length; i++) {
          if (parseInt(exprTree[i][3]).toString() === "NaN") {
            return "#VAL!";
          }
          else if (parseInt(exprTree[i][3]) < 0) {
            return "#NUM!";
          }
        }

        var r;
        var a;
        var result;
        for(i = exprTree.length - 1, result = parseInt(exprTree[i][3]); i;) {
          for(a = parseInt(exprTree[--i][3]); r = a % result; a = result, result = r);
        }
        
        return result;
      }
    ),
    new Array (
      "INT",
      "Rounds a number down to the nearest integer",
      "(number)",
      function (exprTree) {
        if (exprTree.length != 1 || parseInt(exprTree[0][3]).toString() === "NaN") {
          return "#FUNCTION!";
        }

        var arg = parseInt(exprTree[0][3]);

        return parseInt(exprTree[0][3]);
      }
    ),
    new Array (
      "LCM",
      "Returns the least common multiple",
      "(number1, number2, ...)",
      function (exprTree) {
        if (exprTree.length < 2) {
          return "#FUNCTION!";
        }

        // Parameter sanity
        var i;
        var o = new Array();
        for (i = 0; i < exprTree.length; i++) {
          if (parseInt(exprTree[i][3]).toString() === "NaN") {
            return "#VAL!";
          }
          else if (parseInt(exprTree[i][3]) < 0) {
            return "#NUM!";
          }

          o.push(parseInt(exprTree[i][3]));
        }

        for(var i, j, n, d, r = 1; (n = o.pop()) !== undefined;) {
          while(n > 1) {
            if(n % 2) {
              for (i = 3, j = Math.floor(Math.sqrt(n)); i <= j && n % i; i += 2);
              d = i <= j ? i : n;
            }
            else {
              d = 2;
            }
            for(n /= d, r *= d, i = o.length; i; !(o[--i] % d) && (o[i] /= d) == 1 && o.splice(i, 1));
          }
        }

        return r;
      }
    ),
    new Array (
      "LN",
      "Returns the natural logarithm of a number",
      function () { }
    ),
    new Array (
      "LOG",
      "Returns the logarithm of a number to a specified base",
      "(number, [base])",
      function (exprTree) {
        if (exprTree.length != 1 && exprTree.length != 2) {
          return "#FUNCTION!";
        }

        if (parseInt(exprTree[0][3]).toString() == "NaN" || parseInt(exprTree[0][3]) < 0) {
          return "#NUM!";
        }

        var arg = parseInt(exprTree[0][3]);
        var base = 10;
        
        if (exprTree.length == 2) {
          if (parseFloat(exprTree[1][3]).toString() === "NaN") {
            return "#NUM!";
          }

          base = parseFloat(exprTree[1][3]);
        }

        return Math.log(arg) / Math.log(base);
      }
    ),
    new Array (
      "LOG10",
      "Returns the base-10 logarithm of a number",
      "(number)",
      function (exprTree) {
        if (exprTree.length != 1) {
          return "#FUNCTION!";
        }

        if (parseInt(exprTree[0][3]).toString() == "NaN" || parseInt(exprTree[0][3]) < 0) {
          return "#NUM!";
        }

        var arg = parseInt(exprTree[0][3]);
        var base = 10;

        return Math.log(arg) / Math.log(base);
      }
    ),
    new Array (
      "MDETERM",
      "Returns the matrix determinant of an array",
      function () { }
    ),
    new Array (
      "MINVERSE",
      "Returns the matrix inverse of an array",
      function () { }
    ),
    new Array (
      "MMULT",
      "Returns the matrix product of two arrays",
      function () { }
    ),
    new Array (
      "MOD",
      "Returns the remainder from division",
      "(number, divisor)",
      function (exprTree) {
        if (exprTree.length != 2) {
          return "#FUNCTION!";
        }

        var num = parseFloat(exprTree[0][3]);
        var div = parseFloat(exprTree[1][3]);

        if (num.toString() === "NaN" || div.toString() === "NaN") {
          return "#NUM!";
        }

        if (div == 0) {
          return "#DIV/0!";
        }

        var ret = num % div;

        if (div < 0 && ret > 0) {
          ret = 0 - ret;
        }

        return ret;
      }
    ),
    new Array (
      "MROUND",
      "Returns a number rounded to the desired multiple",
      function () { }
    ),
    new Array (
      "MULTINOMIAL",
      "Returns the multinomial of a set of numbers",
      function () { }
    ),
    new Array (
      "ODD",
      "Rounds a number up to the nearest odd integer",
      "(number)",
      function (exprTree) {
        if (exprTree.length != 1) {
          return "#FUNCTION!";
        }

        if (parseInt(exprTree[0][3]).toString() === "NaN") {
          return "#VAL!";
        }

        var neg = false;

        if (parseInt(exprTree[0][3]) < 0) {
          neg = true;
        }

        var ret = Math.round(Math.abs(parseInt(exprTree[0][3])));

        if (ret % 2 == 0) {
          ret++;
        }

        if (neg == true) {
          ret = 0 - ret;
        }

        return ret;
      }
    ),
    new Array (
      "PI",
      "Returns the value of pi",
      function () {
        return parseFloat(Math.PI.toString().substr(0, 16));
      }
    ),
    new Array (
      "POWER",
      "Returns the result of a number raised to a power",
      "(number, power)",
      function (exprTree) {
        if (exprTree.length != 2) {
          return "#FUNCTION!";
        }

        var num1 = parseFloat(exprTree[0][3]);
        var num2 = parseFloat(exprTree[1][3]);

        if (num1.toString() == "NaN" || num2.toString() == "NaN") {
          return "#VAL!";
        }

        return Math.pow(num1, num2);
      }
    ),
    new Array (
      "PRODUCT",
      "Multiplies its arguments",
      function () { }
    ),
    new Array (
      "QUOTIENT",
      "Returns the integer portion of a division",
      function () { }
    ),
    new Array (
      "RADIANS",
      "Converts degrees to radians",
      function () { }
    ),
    new Array (
      "RAND",
      "Returns a random number between 0 and 1",
      function () { }
    ),
    new Array (
      "RANDBETWEEN",
      "Returns a random number between the numbers you specify",
      function () { }
    ),
    new Array (
      "ROMAN",
      "Converts an arabic numeral to roman, as text",
      function () { }
    ),
    new Array (
      "ROUND",
      "Rounds a number to a specified number of digits",
      function () { }
    ),
    new Array (
      "ROUNDDOWN",
      "Rounds a number down, toward zero",
      function () { }
    ),
    new Array (
      "ROUNDUP",
      "Rounds a number up, away from zero",
      function () { }
    ),
    new Array (
      "SERIESSUM",
      "Returns the sum of a power series based on the formula",
      function () { }
    ),
    new Array (
      "SIGN",
      "Returns the sign of a number",
      function () { }
    ),
    new Array (
      "SIN",
      "Returns the sine of the given angle",
      function () { }
    ),
    new Array (
      "SINH",
      "Returns the hyperbolic sine of a number",
      function () { }
    ),
    new Array (
      "SQRT",
      "Returns a positive square root",
      function () { }
    ),
    new Array (
      "SQRTPI",
      "Returns the square root of (number * pi)",
      function () { }
    ),
    new Array (
      "SUBTOTAL",
      "Returns a subtotal in a list or database",
      function () { }
    ),
    new Array (
      "SUM",
      "Adds its arguments",
      "(number1, number2, ...)",
      function (exprTree) {
        result = 0;

        var i;
        for (i = 0; i < exprTree.length && exprTree[i][2] !== "stop"; i++) {
          if (parseFloat(exprTree[i][3]).toString() !== "NaN") {
            result += parseFloat(exprTree[i][3]);
          }
        }

        return result;
      }
    ),
    new Array (
      "SUMIF",
      "Adds the cells specified by a given criteria",
      function () { }
    ),
    new Array (
      "SUMPRODUCT",
      "Returns the sum of the products of corresponding array components",
      function () { }
    ),
    new Array (
      "SUMSQ",
      "Returns the sum of the squares of the arguments",
      function () { }
    ),
    new Array (
      "SUMX2MY2",
      "Returns the sum of the difference of squares of corresponding values in two arrays",
      function () { }
    ),
    new Array (
      "SUMX2PY2",
      "Returns the sum of the sum of squares of corresponding values in two arrays",
      function () { }
    ),
    new Array (
      "SUMXMY2",
      "Returns the sum of squares of differences of corresponding values in two arrays",
      function () { }
    ),
    new Array (
      "TAN",
      "Returns the tangent of a number",
      "(number)",
      function (exprTree) {
        if (exprTree.length != 1 || parseFloat(exprTree[0][3]).toString() === "NaN") {
          return "#FUNCTION!";
        }

        var arg = parseFloat(exprTree[0][3]);

        return Math.tan(arg);
      }
    ),
    new Array (
      "TANH",
      "Returns the hyperbolic tangent of a number",
      "(number)",
      function (exprTree) {
        if (exprTree.length != 1 || parseFloat(exprTree[0][3]).toString() === "NaN") {
          return "#FUNCTION!";
        }

        var arg = parseFloat(exprTree[0][3]);

        return (Math.exp(arg) - Math.exp(-arg)) / (Math.exp(arg) + Math.exp(-arg));
      }
    ),
    new Array (
      "TRUNC",
      "Truncates a number to an integer",
      "(float)",
      function (exprTree) {
        if (exprTree.length != 1 || parseInt(exprTree[0][3]).toString() === "NaN") {
          return "#FUNCTION!";
        }

        return parseInt(exprTree[0][3]);
      }
    ),

    // Text functions
    new Array (
      "ASC",
      "Changes full-width (double-byte) English letters or " +
      "katakana within a character string to half-width (single-byte) characters",
      function () { }
    ),
    new Array (
      "BAHTTEXT",
      "Converts a number to text, using the ß (baht) currency format",
      function () { }
    ),
    new Array (
      "CHAR",
      "Returns the character specified by the code number",
      "(number)",
      function (exprTree) {
        if (exprTree.length != 1 || parseInt(exprTree[0][3]).toString() === "NaN") {
          return "#FUNCTION!";
        }

        return String.fromCharCode(parseInt(exprTree[0][3]));
      }
    ),
    new Array (
      "CLEAN",
      "Removes all nonprintable characters from text",
      function () { }
    ),
    new Array (
      "CODE",
      "Returns a numeric code for the first character in a text string",
      "(text)",
      function (exprTree) {
        if (exprTree.length != 1 || exprTree[0][3].length == 0) {
          return "#FUNCTION!";
        }

        return exprTree[0][3].charCodeAt(0);
      }
    ),
    new Array (
      "CONCATENATE",
      "Joins several text items into one text item",
      "(text1, text2, ...)",
      function (exprTree) {
        result = "";

        var i;
        for (i = 0; i < exprTree.length && exprTree[i][2] !== "stop"; i++) {
          result += exprTree[i][3];
        }

        return result;
      }
    ),
    new Array (
      "DOLLAR",
      "Converts a number to text, using the $ (dollar) currency format",
      function () { }
    ),
    new Array (
      "EXACT",
      "Checks to see if two text values are identical",
      function (exprTree) {
        if (exprTree.length != 2) {
          return "#FUNCTION!";
        }

        return exprTree[0][3] === exprTree[1][3] ? "TRUE" : "FALSE";
      }
    ),
    new Array (
      "FIND, FINDB",
      "Finds one text value within another (case-sensitive)",
      function () { }
    ),
    new Array (
      "FIXED",
      "Formats a number as text with a fixed number of decimals",
      function () { }
    ),
    new Array (
      "JIS",
      "Changes half-width (single-byte) English letters or katakana " +
      "within a character string to full-width (double-byte) characters",
      function () { }
    ),
    new Array (
      "LEFT",
      "Returns the leftmost characters from a text value",
      "(text, totalChars)",
      function (exprTree) {
        if (exprTree.length != 2 || exprTree[0][3].length == 0 || parseInt(exprTree[1][3]).toString() === "NaN" || exprTree[0][3].length < parseInt(exprTree[1][3])) {
          return "#FUNCTION!";
        }

        var len = parseInt(exprTree[1][3]);

        return exprTree[0][3].substr(0, len);
      }
    ),
    new Array (
      "LEN",
      "Returns the number of characters in a text string",
      "(text)",
      function (exprTree) {
        if (exprTree.length != 1) {
          return "#FUNCTION!";
        }

        return exprTree[0][3].length;
      }
    ),
    new Array (
      "LOWER",
      "Converts text to lowercase",
      function (exprTree) {
        if (exprTree.length != 1) {
          return "#FUNCTION!";
        }

        return exprTree[0][3].toLowerCase();
      }
    ),
    new Array (
      "MID, MIDB",
      "Returns a specific number of characters from a text string starting at the position you specify",
      function () { }
    ),
    new Array (
      "PHONETIC",
      "Extracts the phonetic (furigana) characters from a text string",
      function () { }
    ),
    new Array (
      "PROPER",
      "Capitalizes the first letter in each word of a text value",
      function () { }
    ),
    new Array (
      "REPLACE, REPLACEB",
      "Replaces characters within text",
      function () { }
    ),
    new Array (
      "REPT",
      "Repeats text a given number of times",
      function (exprTree) {
        if (exprTree.length != 2 || parseInt(exprTree[1][3]).toString() === "NaN") {
          return "#FUNCTION!";
        }

        var i;
        var result = "";
        for (i = 0; i < parseInt(exprTree[1][3]); i++) {
          result += exprTree[0][3];
        }

        return result;
      }
    ),
    new Array (
      "RIGHT",
      "Returns the rightmost characters from a text value",
      "(text, totalChars)",
      function (exprTree) {
        if (exprTree.length != 2 || exprTree[0][3].length == 0 || parseInt(exprTree[1][3]).toString() === "NaN" || exprTree[0][3].length < parseInt(exprTree[1][3])) {
          return "#FUNCTION!";
        }

        var len = parseInt(exprTree[1][3]);

        return exprTree[0][3].substr(exprTree[0][3].length - len);
      }
    ),
    new Array (
      "SEARCH, SEARCHB",
      "Finds one text value within another (not case-sensitive)",
      function () { }
    ),
    new Array (
      "SUBSTITUTE",
      "Substitutes new text for old text in a text string",
      function () { }
    ),
    new Array (
      "T",
      "Converts its arguments to text",
      function () { }
    ),
    new Array (
      "TEXT",
      "Formats a number and converts it to text",
      function () { }
    ),
		new Array (
			"TRIM",
			"Removes spaces from text",
			"(expression)",
			function (exprTree) {
				if (exprTree.length != 1) {
					return "#FUNCTION!";
				}

				return exprTree[0][3].trim();
			}
		),
		new Array (
			"UPPER",
			"Converts text to uppercase",
			"(expression)",
			function (exprTree) {
				if (exprTree.length != 1) {
					return "#FUNCTION!";
				}

				return exprTree[0][3].toUpperCase();
			}
		),
    new Array (
      "VALUE",
      "Converts a text argument to a number",
      function () { }
    ),

    // Information functions
    new Array (
      "CELL",
      "Returns information about the formatting, location, or contents of a cell",
      function () { }
    ),
    new Array (
      "ERROR.TYPE",
      "Returns a number corresponding to an error type",
      function () { }
    ),
    new Array (
      "INFO",
      "Returns information about the current operating environment",
      function () { }
    ),
    new Array (
      "ISBLANK",
      "Returns TRUE if the value is blank",
      function () { }
    ),
    new Array (
      "ISERR",
      "Returns TRUE if the value is any error value except #N/A",
      function () { }
    ),
    new Array (
      "ISERROR",
      "Returns TRUE if the value is any error value",
      function () { }
    ),
    new Array (
      "ISEVEN",
      "Returns TRUE if the number is even",
      function () { }
    ),
    new Array (
      "ISLOGICAL",
      "Returns TRUE if the value is a logical value",
      function () { }
    ),
    new Array (
      "ISNA",
      "Returns TRUE if the value is the #N/A error value",
      function () { }
    ),
    new Array (
      "ISNONTEXT",
      "Returns TRUE if the value is not text",
      function () { }
    ),
    new Array (
      "ISNUMBER",
      "Returns TRUE if the value is a number",
      function () { }
    ),
    new Array (
      "ISODD",
      "Returns TRUE if the number is odd",
      function () { }
    ),
    new Array (
      "ISREF",
      "Returns TRUE if the value is a reference",
      function () { }
    ),
    new Array (
      "ISTEXT",
      "Returns TRUE if the value is text",
      function () { }
    ),
    new Array (
      "N",
      "Returns a value converted to a number",
      function () { }
    ),
    new Array (
      "NA",
      "Returns the error value #N/A",
      function () { }
    ),
    new Array (
      "TYPE",
      "Returns a number indicating the data type of a value",
      function () { }
    )
  );

	this.supportedExcelFunctions.sort (
		function (a, b) {
			var tmpa = a[0];
			var tmpb = b[0];

			tmpa = tmpa.toLowerCase();
			tmpb = tmpb.toLowerCase();

			if (tmpa > tmpb)
				return 1;
			if (tmpa < tmpb)
				return -1;

			return 0;
		}
	);
  

	this.getArgumentType = function (val) {
		// For now the only types supported are "number" and "string"

		if (parseFloat(val).toString() === "NaN")
			return "string";

		var v1 = parseFloat(val);
		var v2 = parseInt(val);

		if (v1 === v2)
			return "integer";
		else
			return "float";

		return "number";
	}

	this.parseFunction = function (exprTree) {
		var result = 0;

		// Build a new parse tree that contains only
		// the arguments (and any ranges expanded to their values)
		// For the purposes of calling the respective function
		// we only care about type "operand". We also want to ignore
		// any operand delimiters
		var newParseTree = new Array();
		var i;

		for (i = 0; i < exprTree.length; i++) {
			if (exprTree[i] === undefined)
				return "#ERROR!";

			if (exprTree[i][1] == "operand") {
				// Ensure that this is not an operand delimiter
				if (exprTree[i][3] !== ",") {
					// Is this a range or a value?
					if (exprTree[i][2] == "range") {
						// Translate the range into their respective values and push them onto the stack

						// First we need to see if this is an actual range or just a singular cell reference
						var tmp = exprTree[i][3];
						var tmpArrOp = new Array();

						if (tmp.match(/:/)) {
							// This is an actual range
							var tops = tmp.split(/:/);

							if (tops.length != 2)
								return "#NAME?";

							var tmpColRef = new Array();
							tmpColRef.push("");
							tmpColRef.push("");

							var cols = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
							var m;
							for (m = 0; m < tops[0].length; m++) {
								if (cols.search(tops[0].charAt(m).toUpperCase()) !== -1)
									tmpColRef[0] += tops[0].charAt(m).toUpperCase();
							}

							for (m = 0; m < tops[1].length; m++) {
								if (cols.search(tops[1].charAt(m).toUpperCase()) !== -1)
									tmpColRef[1] += tops[1].charAt(m).toUpperCase();
							}

							var start = this.getCellReferenceIndex(tops[0]);
							var end = this.getCellReferenceIndex(tops[1]);
							if ((tmpColRef[0] !== tmpColRef[1] && tops[0].substr(tmpColRef[0].length) !== tops[1].substr(tmpColRef[1].length))
									|| (parseInt(tops[0].substr(tmpColRef[0].length)) > this.rows[this.selectedSheet][1].length)
									|| (parseInt(tops[1].substr(tmpColRef[1].length)) > this.rows[this.selectedSheet][1].length)
									|| (start > this.rows[this.selectedSheet][1][0].length)
									|| (end > this.rows[this.selectedSheet][1][0].length)
									|| (parseInt(tops[0].substr(tmpColRef[0].length)) <= 0)
									|| (parseInt(tops[1].substr(tmpColRef[1].length)) <= 0)) {
								return "#REF!";
							}

							if (start === end) {
								var col = this.getCellReferenceVal(tops[0]);
								var start = this.getRowReferenceIndex(tops[0]);
								var end = this.getRowReferenceIndex(tops[1]);

								var k;
								for (k = start; k <= end; k++) {
									var tmpVal = this.translateCellReference(col + (k + 1));

									var arr = new Array (
										1,
										"operand",
										this.getArgumentType(tmpVal),
										tmpVal
									);
                    
									newParseTree.push(arr);
								}
							}
							else {
								var k;
								for (k = start; k <= end; k++) {
									var cols = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
									var tmpVal = this.translateCellReference(cols[k] + tops[0].substr(1));

									var arr = new Array (
										1,
										"operand",
										this.getArgumentType(tmpVal),
										tmpVal
									);
                    
									newParseTree.push(arr);
								}
							}
						}
						else {
							// cell reference
							var tmpVal = this.translateCellReference(tmp);
							if (tmpVal === undefined)
								tmpVal = tmp;

							var arr = new Array (
								1,
								"operand",
								this.getArgumentType(tmpVal),
								tmpVal
							);

							newParseTree.push(arr);
						}
					}
					else {
						// No translation necessary, just push onto the stack
						newParseTree.push(exprTree[i]);
					}
				}
			}
		}

		// Now call the respective function with arguments newParseTree
		for (i = 0; i < this.supportedExcelFunctions.length; i++) {
			if (this.supportedExcelFunctions[i][0].toUpperCase() == exprTree[0][3].toUpperCase()) {
				var ret = this.supportedExcelFunctions[i][this.supportedExcelFunctions[i].length - 1](newParseTree);

				if (ret === 0)
					ret = "0";
				return ret;
			}
		}
	}

	this.computeSubExpression = function (exprTree) {
		var ret = exprTree;

		var i;
		var operands = new Array();
		var action = 0;
		var results = 0;

		var str = "";
		for (i = 0; i < ret.length; i++) {
			var index = ret[i][0];
			var type = ret[i][1];
			var subtype = ret[i][2];
			var tokens = ret[i][3];
			str += "  Index: " + ret[i][0] + "\n";
			str += "   Type: " + ret[i][1] + "\n";
			str += "Subtype: " + ret[i][2] + "\n";
			str += " Tokens: " + ret[i][3] + "\n";
			str += "\n";

			switch (type) {
				case "operand":
					if (subtype == "number") {
						operands.push(tokens);
					}
					else if (subtype == "range") {
						var tmp = this.translateCellReference(tokens);
						if (tmp === undefined)
							operands.push(tokens);
						else
							operands.push(tmp);
					}
					break;
				case "operator-infix":
					if ((subtype == "math") || (subtype == "concatenate"))
						operands.push(tokens);
					break;
				case "operator-postfix":
					if ((subtype == "") && (tokens == "%"))
						operands.push(tokens);
					break;
				default:
					break;
			}
		}

		var debugDiv = document.getElementById("BNGridDebug");
		if (debugDiv !== undefined)
			debugDiv.innerHTML = "<pre>" + str + "</pre>";

		if (operands.length == 1)
			return operands[0];

		str = "";
		var result = results;
		var op = 0;
		for (i = 0; i < operands.length; i++) {
			switch (operands[i]) {
				case "+":
					i++;

					if (parseFloat(operands[i]).toString() !== "NaN")
						result += parseFloat(operands[i]);
					break;
				case "-":
					i++;

					if (parseFloat(operands[i]).toString() !== "NaN")
						result -= parseFloat(operands[i]);
					break;
				case "*":
					i++;

					if (parseFloat(operands[i]).toString() !== "NaN")
						result *= parseFloat(operands[i]);
					break;
				case "/":
					i++;

					if (parseFloat(operands[i]).toString() !== "NaN")
						result /= parseFloat(operands[i]);
					break;
				case "%":
					i++;

					if (parseFloat(operands[i]).toString() !== "NaN")
						result %= parseFloat(operands[i]);
					break;
				case "&":
					i++;
					result += operands[i].toString();
					break;
				default:
					if (parseFloat(operands[i]).toString() !== "NaN")
						result = parseFloat(operands[i]);
					else
						result = operands[i];
					break;
			}
		}

		return result;
	}

	// =1 * (((2 + 3) * 4) + 5)
	// =B1 * (((B2 + B3) * C1) + C2)
	this.parseSubExpressions = function (parseTree) {
		var results = 0;
		var ptOps = this.getFunction(parseTree);
		while (ptOps !== undefined) {
			results = this.parseFunction(ptOps);

			var str = "";
			for (i = 0; i < ptOps.length; i++) {
				if (ptOps[i] === undefined)
					return "#ERROR!";

				var index = ptOps[i][0];
				var type = ptOps[i][1];
				var subtype = ptOps[i][2];
				var tokens = ptOps[i][3];
				str += "  Index: " + index + "\n";
				str += "   Type: " + type + "\n";
				str += "Subtype: " + subtype + "\n";
				str += " Tokens: " + tokens + "\n";
				str += "\n";
			}

			// Now remove this from the parse tree
			var newParseTree = new Array();
			var i;
			var included = 0;
			for (i = 0; i < parseTree.length; i++) {
				var j;
				var okToAdd = true;
				for (j = 0; j < ptOps.length; j++) {
					if (parseTree[i][0] == ptOps[j][0])
						okToAdd = false;
				}

				if (okToAdd == true) {
					newParseTree.push(parseTree[i]);
				}
				else {
					if (included == 0) {
						var ar = new Array();
						ar.push(parseTree[i][0]);
						ar.push("operand");
						ar.push("number");
						ar.push(results);

						newParseTree.push(ar);
						included++;
					}
				}
			}

			parseTree = newParseTree;
			str += "=======================================" + "\n";
			for (i = 0; i < parseTree.length; i++) {
				var index = parseTree[i][0];
				var type = parseTree[i][1];
				var subtype = parseTree[i][2];
				var tokens = parseTree[i][3];
				str += "  Index: " + index + "\n";
				str += "   Type: " + type + "\n";
				str += "Subtype: " + subtype + "\n";
				str += " Tokens: " + tokens + "\n";
				str += "\n";
			}

			var debugDiv = document.getElementById("BNGridDebug");
			if (debugDiv !== undefined)
				debugDiv.innerHTML = "<pre>" + str + "</pre>";

			ptOps = this.getFunction(parseTree);
		}

		ptOps = this.getInnerSubExpression(parseTree);
		while (ptOps.length > 0) {
			// Calclulate the value then remove this parse tree element
			var result = this.computeSubExpression(ptOps);

			// Now remove this from the parse tree
			var newParseTree = new Array();
			var i;
			var included = 0;
			for (i = 0; i < parseTree.length; i++) {
				var j;
				var okToAdd = true;
				for (j = 0; j < ptOps.length; j++) {
					if (parseTree[i][0] == ptOps[j][0])
						okToAdd = false;
				}

				if (okToAdd == true) {
					newParseTree.push(parseTree[i]);
				}
				else {
					if (included == 0) {
						var ar = new Array();
						ar.push(parseTree[i][0]);
						ar.push("operand");
						ar.push("number");
						ar.push(result);

						newParseTree.push(ar);
						included++;
					}
				}
			}

			parseTree = newParseTree;
			ptOps = this.getInnerSubExpression(parseTree);
		}

		return parseTree;
	}

	// =1 * (((2 + 3) * 4) + 5)
	this.parseExcelFormula = function (value) {
		var ret = parseFormula(value);
		var str = "";

		var i;
		for (i = 0; i < ret.length; i++) {
			var index = ret[i][0];
			var type = ret[i][1];
			var subtype = ret[i][2];
			var tokens = ret[i][3];
			str += "  Index: " + ret[i][0] + "\n";
			str += "   Type: " + ret[i][1] + "\n";
			str += "Subtype: " + ret[i][2] + "\n";
			str += " Tokens: " + ret[i][3] + "\n";
			str += "\n";

			// Ensure that negative numbers are preserved
			if (i > 0) {
				if ((ret[i - 1][1] == "operator-prefix") && (ret[i - 1][3] == "-") && (ret[i][2] == "number"))
					ret[i][3] = "-" + ret[i][3];
			}
		}

		ret = this.parseSubExpressions(ret);

		var fin = this.computeSubExpression(ret);
		if ((fin === undefined) || (fin.toString() === "NaN"))
			return "#FORMULA!";

		var tfin = this.parseFunction(ret);
		if (tfin !== undefined)
			fin = tfin;

		var debugDiv = document.getElementById("BNGridDebug");
		if (debugDiv !== undefined)
			debugDiv.innerHTML = "<pre>" + str + "</pre>";

		return fin;
	}

	this.highlightCell = function (cell, _this) {
		var rowcol = cell.getAttribute("id");

		var tmp = rowcol.split(";");

		tmp[0] = parseInt(tmp[0].replace("row", "")) - 1;
		tmp[1] = parseInt(tmp[1].replace("col", "")) - 1;

		var border = "border: 1px solid #000000;";
		var rows = _this.getRows(_this.selectedSheet);
		if (tmp[0] == 0)
			border += " border-top: 0px;";
		else if (tmp[0] == (rows.length - 1))
			border += " border-bottom: 0px;";

		if (tmp[1] == (rows[0].length - 1))
			border += " border-right: 0px;";

		cell.setAttribute (
			"style",
			border +
			" background-color: #EEEEFF;"
		);
	}

	this.highlightHeader = function (val) {
		var i;
		var rows = this.rows[this.selectedSheet][1];
		for (i = 0; i < rows.length; i++) {
			var el = document.getElementById("rowh" + (i + 1));
			el.setAttribute("style", "background-color: #e2ebf2");

			var j;
			for (j = 0; j < rows[i].length; j++) {
				var el = document.getElementById("colh" + (j + 1));
				el.setAttribute("style", "background-color: #e2ebf2");

				el = document.getElementById("row" + (i + 1) + ";col" + (j + 1));
				el.setAttribute("style", "background-color: #FFFFFF");
			}
		}

		var arr = val.split(";");
		var row = arr[0];
		var col = arr[1];

		row = row.replace(/^row/, "");
		col = col.replace(/^col/, "");

		var el = document.getElementById("colh" + col);
		el.setAttribute("style", "background: url(); background-color: #fbc271;");

		el = document.getElementById("rowh" + row);
		el.setAttribute("style", "background-color: #fbc271;");
	}

	this.showPendingFunctions = function () {
		var i;
		var ct = "<pre>";
		for (i = 0; i < this.supportedExcelFunctions.length; i++) {
			if (this.supportedExcelFunctions[i].length == 3)
				ct += this.supportedExcelFunctions[i][0] + "\n";
		}
		ct += "</pre>";

		var id = document.getElementById("BNGridDebug");
		id.innerHTML = ct;
	}

	this.selectCell = function (cell) {
		this.highlightHeader(cell);
		this.highlightCell(document.getElementById(cell), this);

		var id = document.getElementById("formulaEditBox");
		var tmp = cell.split(";");

		tmp[0] = tmp[0].replace("row", "");
		tmp[1] = tmp[1].replace("col", "");

		tmp[0] = parseInt(tmp[0]) - 1;
		tmp[1] = parseInt(tmp[1]) - 1;

		var formula = this.getFormula(tmp[0], tmp[1]);
		if (formula.length == 0) {
			var val = document.getElementById(cell).innerHTML;
			if (val == "&nbsp;")
				val = "";
			id.setAttribute("value", val);
		}
		else
			id.setAttribute("value", formula);

		this.selectedCell = cell;
		var id = document.getElementById("formulaEditBox");
		id.focus();
		id.select();
	}

	this.editCell = function(cell, _this) {
		if (document.getElementById) {
			_this.editing = true;
			cell.normalize();
			var _this = this;
			var input = document.createElement("input");
			input.setAttribute("style", "width: 100%; border: 1px solid #000000;");
			input.setAttribute("id", "celledit");

			var id = cell.getAttribute("id");

			var arr = id.split(";");
			var row = arr[0];
			var col = arr[1];

			row = row.replace(/^row/, "");
			col = col.replace(/^col/, "");

			row = parseInt(row) - 1;
			col = parseInt(col) - 1;

			var val = this.getFormula(row, col);
			if (val !== "") {
				input.setAttribute("value", val);
			}
			else {
				if (cell.firstChild) {
					input.setAttribute("value", cell.firstChild.nodeValue);
				}
			}

			input.onchange = function (evt) {
				_this.setCell(this.parentNode, this.value);
				_this.editing = false;
			};

			input.onclick = function (evt) {
				evt.cancelBubble = true;
				if (evt.stopPropagation) {
					evt.stopPropagation();
				}
			};

			input.onkeypress = function (e) {
				var id = document.getElementById("tooltip");
				if (id === undefined)
					return;

				if (e.keyCode == 13) {
					this.onblur();
					return;
				}

				var val = this.value;
				if (!val.match(/^=/))
					return;

				id.style.top = event.clientY;
				id.style.left = event.clientX + 10;

				var oldVal = val;

				val = val.replace(/^=/, "");
				val = val.replace(/\(.*/, "");
				val = val.toUpperCase();

				// val now holds our potential match
				var i;
				var funcRef = "";
				var ignoreFunction = "";
				for (i = 0; i < _this.supportedExcelFunctions.length; i++) {
					var funcName = _this.supportedExcelFunctions[i][0].toUpperCase();
					if (val.length <= funcName.length) {
						if (funcName == val) {
							// Looking for an exact match
							if (funcName == val) {
								ignoreFunction = funcName;
								funcRef = "<strong>" + _this.supportedExcelFunctions[i][0] + "</strong>";

								if (_this.supportedExcelFunctions[i].length > 3) {
									var args = _this.supportedExcelFunctions[i][2].split(/,/);
									var totalArgs = args.length;
									var j;
									var entered = oldVal.split(",");
									var enteredArgs = entered.length;

									for (j = 0; j < totalArgs; j++) {
//										if ((j == (enteredArgs - 1) || ((enteredArgs > j) && (j == (totalArgs - 1)))) && oldVal.match(/\(/))
//											funcRef += "<strong>";
										funcRef += args[j];
									}

//									if ((j == (enteredArgs - 1) || (enteredArgs > j && j == (totalArgs - 1))) && oldVal.match(/\(/))
//										funcRef += "</strong>";
									if (j < (totalArgs - 1))
										funcRef += ",";
								}
							}
							else {
								funcRef += "()";
							}

							funcRef += "<div style=\"width: 100%; border-top: 1px solid #CCCCCC;\" />";
							funcRef += _this.supportedExcelFunctions[i][1];
						}
					}
				}

				var maxTopMatches = 10;
				var topMatches = new Array();
				var searchLength = val.length;
				for (i = 0; i < _this.supportedExcelFunctions.length; i++) {
					var funcName = _this.supportedExcelFunctions[i][0].toUpperCase();
					if (funcName.substr(0, searchLength) == val) {
						if (topMatches.length < maxTopMatches && _this.supportedExcelFunctions[i][0] !== ignoreFunction)
							topMatches.push(_this.supportedExcelFunctions[i]);
					}

					if ((i == (_this.supportedExcelFunctions.length - 1)) && (topMatches.length < maxTopMatches) && ((searchLength - 1) > 0)) {
						i = 0;
						searchLength--;
					}
				}

				if (ignoreFunction !== "" && topMatches.length > 0) {
					funcRef += "<br><br>";
					funcRef += "<em>Similar functions</em>:<br>";
				}

				// Now add the top 5 to the results
				if (topMatches.length > 0)
					funcRef += "<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"padding: 0px;\">";

				for (i = 0; i < topMatches.length; i++) {
					funcRef += "<tr>";
					funcRef += "<td style=\"padding-right: 4px;\">" + topMatches[i][0] + "</td>";
					funcRef += "<td>" + topMatches[i][1] + "</td>";
					funcRef += "</tr>";

					if (i != (topMatches.length - 1))
						funcRef += "<tr><td colspan=\"2\" style=\"border-bottom: 1px solid #EEEEEE;\"></td></tr>";
				}

				if (topMatches.length > 0)
					funcRef += "</table>";

				if (funcRef.length > 0) {
					id.innerHTML = funcRef;
					id.style.display = "block";
				}

				return true;
			}

			input.onblur = input.onchange;

			if (cell.firstChild)
				cell.replaceChild(input, cell.firstChild);
			else
				cell.appendChild(input);

			var div = document.createElement("div");
			div.setAttribute("id", "tooltip");
			div.setAttribute (
				"style",
				"position: absolute; " +
				"padding: 4px; " +
				"display: none; " +
				"background-color: #FFFFFF; " +
				"border: 1px solid #000000;"
			);
			div.innerHTML = "";
			cell.appendChild(div);

			input.focus();
			input.select();
		}
		else if (document.all) {
			cell.innerHTML =
				"<INPUT " +
				" id=\"editCell\"" +
				" onclick=\"event.cancelBubble = true;\"" + 
				" ondblclick=\"event.cancelBubble = true;\"" + 
				" onchange=\"setCell(this.parentElement, this.value);\" " +
				" onblur=\"setCell(this.parentElement, this.value);\" " +
				" value=\"" + cell.innerText + "\"" +
				" size=\"" + cell.innerText.length + "\"" +
				">";
			document.all.editCell.focus();
			document.all.editCell.select();
		}
	}

	this._removeElement = function (elId) {
		if (elId !== undefined && elId !== null) {
			var el = document.getElementById(elId);

			if (el !== undefined && el !== null) {
				var par = el.parentNode;
				par.removeChild(el);
			}
		}
	}

	this.addColumn = function () {
		var rows = this.getRows(this.selectedSheet);
		if (rows.length == 0) {
			this.add("");
			this.draw();
			return;
		}

		var after = rows[0].length;
		if (this.selectedColumn != -1)
			after = this.selectedColumn;
		after--;

		var i;
		for (i = 0; i < rows.length; i++) {
			var cols = rows[i];
			var newCols = new Array();

			var j;
			for (j = 0; j < cols.length; j++) {
				newCols.push(cols[j]);

				if (j == after)
					newCols.push("&nbsp;");
			}

			rows[i] = newCols;
		}

		this.selectedColumn = -1;

		this.draw();

		var objDiv = document.getElementById("BNGridDataArea");
		objDiv.scrollLeft = objDiv.scrollWidth;
	}

	this.addRow = function () {
		var rows = this.getRows(this.selectedSheet);

		if (rows.length == 0) {
			this.add("");
			this.draw();
			return;
		}

		var arr = rows[0];
		var newArr = new Array();
		var newForm = new Array();
		var i;

		var rows = this.getRows(this.selectedSheet);
		var formulas = this.getFormulas(this.selectedSheet);
		var after = rows.length;
		if (this.selectedRow != -1)
			after = this.selectedRow;
		after--;

		var newRows = new Array();
		var newFormulas = new Array();

		for (i = 0; i < arr.length; i++) {
			newArr.push("&nbsp");
			newForm.push("");
		}

		for (i = 0; i < rows.length; i++) {
			newRows.push(rows[i]);
			newFormulas.push(formulas[i]);
			if (i == (after)) {
				newRows.push(newArr);
				newFormulas.push(newForm);
			}
		}

		this.selectedRow = -1;

		this.rows[this.selectedSheet][1] = newRows;
		this.rows[this.selectedSheet][2] = newFormulas;
		this.draw();

		var objDiv = document.getElementById("BNGridDataArea");
		objDiv.scrollTop = objDiv.scrollHeight;
	}

	this.delColFormula = function () {
		var rows = this.getFormulas(this.selectedSheet);
		var col = rows[0].length;
		if (this.selectedColumn != -1)
			col = this.selectedColumn;
		col--;

		if (rows[0].length > 1) {
			var i;
			for (i = 0; i < rows.length; i++) {
				var j;
				var newCol = new Array();
				for (j = 0; j < rows[i].length; j++) {
					if (j != col)
						newCol.push(rows[i][j]);
				}

				rows[i] = newCol;
			}

			this.selectedColumn = -1;
		}
	}

	this.delColumn = function () {
		var rows = this.getRows(this.selectedSheet);
		var col = rows[0].length;
    	if (this.selectedColumn != -1)
			col = this.selectedColumn;
		col--;

		if (rows[0].length > 1) {
			var i;
			for (i = 0; i < rows.length; i++) {
				var j;
				var newCol = new Array();
				for (j = 0; j < rows[i].length; j++) {
					if (j != col)
						newCol.push(rows[i][j]);
				}

				rows[i] = newCol;
			}

			this.delColFormula();
			this.selectedColumn = -1;
			this.draw();
		}
		else {
			alert("You cannot remove the last column");
		}
	}

	this.delRowFormula = function () {
		var rows = this.getFormulas(this.selectedSheet);
		var row = rows.length;
		if (this.selectedRow != -1)
			row = this.selectedRow;
		row--;

		if (rows.length > 1) {
			var i;
			var newArr = new Array();
			for (i = 0; i < rows.length; i++) {
				if (i != row)
					newArr.push(rows[i]);
			}

			this.selectedRow = -1;
			this.rows[this.selectedSheet][2] = newArr;
		}
	}

	this.delRow = function () {
		var rows = this.getRows(this.selectedSheet);
		var row = rows.length;
		if (this.selectedRow != -1)
			row = this.selectedRow;
		row--;

		if (rows.length > 1) {
			var i;
			var newArr = new Array();
			for (i = 0; i < rows.length; i++) {
				if (i != row)
					newArr.push(rows[i]);
			}

			this.delRowFormula();
			this.selectedRow = -1;
			this.rows[this.selectedSheet][1] = newArr;
			this.draw();
		}
		else {
			alert("You cannot remove the last row");
		}
	}

	this.saveData = function () {
		var par = document.getElementById("BNGridContent");
		var i;

		var sheet;
		for (sheet = 0; sheet < this.rows.length; sheet++) {
			var rows = this.getRows(sheet);
			for (i = 0; i < rows.length; i++) {
				var j;
				for (j = 0; j < rows[i].length; j++) {
					var el = document.createElement("input");
					el.setAttribute("type", "hidden");
					el.setAttribute("name", this.rows[sheet][0] + ";row" + (i + 1) + ";col" + (j + 1));

					var val = rows[i][j];
					if (val == "&nbsp;")
						val = "";
					el.setAttribute("value", val);
					par.appendChild(el);
				}
			}
		}
	}

	this.navigateWorksheet = function (event, form) {
		var key = event.keyCode;
		alert("keyCode: " + key);

		var inpElement = (event.target || event.srcElement);

		var tmp = inpElement.getAttribute("id").split(";");
		if (tmp.length != 2)
			return true;

		tmp[0] = tmp[0].replace(/^row/, "");
		tmp[1] = tmp[1].replace(/^col/, "");

		var row = parseInt(tmp[0]);
		var col = parseInt(tmp[1]);

		if (key == 38 && row > 1) {
			// if the up arrow has been pressed and we're not in the first row
			document.getElementById("row" + (row - 1) + ";col" + col).click;
			return false;
		}
		else if ( key == 40 && row != 2) {
			// if the down arrow has been pressed and we're not in the last row
			return false;
		}
		else if ( key == 37 && col != 1) {
			// if the left arrow has been pressed and we're not in the leftmost column
			return false;
		}
		else if ( key == 39 && col != 8) {
			// if the right arrow has been pressed and we're not in the rightmost column
			return false;
		}

		return true;
	}

	this.generateMenu = function (name, items) {
		var mnDiv = document.createElement(mnDiv);
		mnDiv.setAttribute (
			"style",
			"border: 1px solid #000000; " +
			"position: absolute; " +
			"background-color: #FFFFFF; " +
			"min-width: 40px; " +
			"display: none;"
		);

		mnDiv.onmouseout = function () {
			this.style.display = "none";
		}

		var i;
		for (i = 0; i < items.length; i++) {
			var de = document.createElement("div");
			de.setAttribute("id", "Menu." + name + "." + items[i]);

			de.onmouseover = function () {
				this.setAttribute("style", "background-color: #AAAAFF;");
			}

			de.onmouseout = function () {
				this.setAttribute("style", "background-color: #FFFFFF;");
			}

			if (items[i] == "-") {
				var tmp = document.createElement("div");
				tmp.setAttribute("style", "border-top: 1px solid #000000;");
				de.appendChild(tmp);
			}
			else {
				var txt = document.createTextNode(items[i]);
				de.appendChild(txt);
			}

			mnDiv.appendChild(de);
		}

		return mnDiv;
	}

	this.draw = function () {
		var i;
		var rows = this.getRows(this.selectedSheet);

		var table = document.createElement("table");
		table.setAttribute("class", "gridTable");
		table.setAttribute("cellpadding", "0");
		table.setAttribute("cellspacing", "0");
		table.setAttribute("style", "width: 100%;");

		var thead = document.createElement("thead");
		table.appendChild(thead);

		var tr = document.createElement("tr");
		tr.setAttribute("style", "border: 1px solid #000000;");
		thead.appendChild(tr);

		var alphaChars = " ABCDEFGHIJKLMNOPQRSTUVWXYZ";
		var currIndexCntr = 0;
		var currLetter = " ";

		var colLength = 0;

		if (rows.length > 0)
			colLength = rows[0].length;

		for (i = -1; i < colLength; i++) {
			var border = "";
			if (i < (colLength - 1))
				border += "";

			var index = -1;
			if (i == -1) {
				var th = document.createElement("th");
				th.setAttribute("class", "gridHeader");
				th.setAttribute (
					"style",
					"border-left: 0px; " +
					"background-color: #abc2eb; " +
					"min-width: 15px;"
				);
				th.innerHTML = "&nbsp;";

				tr.appendChild(th);
			}
			else {
				index = alphaChars.search(currLetter);
				if (index + 1 >= alphaChars.length) {
					index = 0;
					currIndexCntr++;
				}

				currLetter = alphaChars.charAt(index + 1);

				var colHead = currLetter;

				if (currIndexCntr > 0)
					colHead = alphaChars.charAt(currIndexCntr) + currLetter;

				var _this = this;
				var newlink = document.createElement("a");
				newlink.setAttribute("href", "#");
				newlink.setAttribute("id", (i + 1));
				newlink.onclick = function() {
					var id = document.getElementById("formulaEditBox");
					id.setAttribute("value", "");
					this.selectedCell = undefined;
					_this.highlightColumn(this.getAttribute("id"));
				};
				newlink.setAttribute("style", "display: block; text-decoration: none; color: #000000;");
				newlink.setAttribute("title", "Select this column");

				var txtCell = document.createTextNode(colHead);
				newlink.appendChild(txtCell);

				var th = document.createElement("th");
				th.setAttribute("class", "gridHeader");
				th.setAttribute("style", border);
				th.setAttribute("id", "colh" + (i + 1));
				th.appendChild(newlink);

				tr.appendChild(th);
			}
		}

		// Now print the actual data
		var tbody = document.createElement("tbody");
		table.appendChild(tbody);
 
		for (i = 0; i < rows.length; i++) {
			var tr = document.createElement("tr");
			tr.setAttribute("id", "row" + (i + 1));
			tr.setAttribute("style", "border: 1px solid #000000;");
			tbody.appendChild(tr);

			var td = document.createElement("td");
			td.setAttribute("class", "gridLeft");
			td.setAttribute("id", "rowh" + (i + 1));

			tr.appendChild(td);

			var newlink = document.createElement("a");
			newlink.setAttribute("href", "#");
			newlink.setAttribute("id", i + 1);
			newlink.onclick = function() {
				var id = document.getElementById("formulaEditBox");
				id.setAttribute("value", "");
				this.selectedCell = undefined;
				_this.highlightRow(this.getAttribute("id"))
			};
			newlink.setAttribute("style", "display: block; text-decoration: none; color: #000000;");
			newlink.setAttribute("title", "Select this row");

			var txtCell = document.createTextNode((i + 1));
			newlink.appendChild(txtCell);

			td.appendChild(newlink);

			if (i == 0)
				td.setAttribute("style", "border-top: 0px;");

			var j;
			for (j = 0; j < rows[i].length; j++) {
				var formula = this.getFormula(i, j);

				if (formula.match(/^=/)) {
					var ret = this.parseExcelFormula(formula);
					rows[i][j] = ret;
				}

				var border = "";
				if (j == 0)
					border += " border-left: 1px solid #000000;";

				var td = document.createElement("td");
				td.setAttribute("id", "row" + (i + 1) + ";col" + (j + 1));
				td.setAttribute("class", "gridEntry");
				td.setAttribute("style", border);

				var _this = this;
				td.onclick = function (event) {
					_this.highlightHeader(this.getAttribute("id"));
					_this.highlightCell(this, _this);

					var id = document.getElementById("formulaEditBox");
					var tmp = this.id.split(";");

					tmp[0] = tmp[0].replace("row", "");
					tmp[1] = tmp[1].replace("col", "");

					tmp[0] = parseInt(tmp[0]) - 1;
					tmp[1] = parseInt(tmp[1]) - 1;

					var formula = _this.getFormula(tmp[0], tmp[1]);
					if (formula.length == 0) {
						var val = this.innerHTML;
						if (val == "&nbsp;")
							val = "";
						id.setAttribute("value", val);
					}
					else
						id.setAttribute("value", formula);

					_this.selectedCell = this.id;
					var id = document.getElementById("formulaEditBox");
					id.focus();
					id.select();
				};

				td.onkeypress = function (e) {
					if (_this.editing == true)
						return;

					_this.highlightHeader(this.getAttribute("id"));
					_this.highlightCell(this, _this);

					var id = document.getElementById("formulaEditBox");
					var tmp = this.id.split(";");

					tmp[0] = tmp[0].replace("row", "");
					tmp[1] = tmp[1].replace("col", "");

					tmp[0] = parseInt(tmp[0]) - 1;
					tmp[1] = parseInt(tmp[1]) - 1;

					var formula = _this.getFormula(tmp[0], tmp[1]);

					if (formula.length == 0)
						id.setAttribute("value", this.innerHTML);
					else
						id.setAttribute("value", formula);

					_this.selectedCell = this.id;
					var id = document.getElementById("formulaEditBox");
					id.focus();
					id.select();
				};

				td.ondblclick = function (e) {
					_this.editCell(this, _this);
					_this.highlightHeader(this.getAttribute("id"));
				};

				tr.appendChild(td);

				var val = rows[i][j];
				if (val == "")
					val = "&nbsp;";
				td.innerHTML = val;
			}
		}

		var _this = this;

		var ct = document.getElementById("BNGridContent");
		ct.innerHTML = "";

		// Menu
		var div = document.createElement("div");
		div.setAttribute (
			"style",
			"background: url(images/background.jpg); " +
			"display: block;"
		);
		div.setAttribute("id", "menu");

		var mnDiv = document.createElement("div");
		var txt = document.createTextNode("File");

		mnDiv.appendChild(txt);

		try {
			mnDiv.style.top = div.clientY;
			mnDiv.style.left = div.clientX;
		} catch (ex) {
			mnDiv.style.top = window.event.clientY;
			mnDiv.style.left = window.event.clientX;
		}

		var fileMenu = this.generateMenu (
			"File",
			new Array (
				"New",
				"Open",
				"Save",
				"-",
				"Exit"
			)
		);

//		mnDiv.appendChild(fileMenu);
//		div.appendChild(mnDiv);
		ct.appendChild(div);

		fileMenu.style.top = "22px";
		fileMenu.style.left = "8px";

		mnDiv.onmouseover = function () {
			fileMenu.style.display = "block";
		}

		// TODO: generate menu

		// Toolbar
		var div = document.createElement("div");
		div.setAttribute (
			"style",
			"background: url(images/background.jpg); " +
			"display: block;"
		);

		div.setAttribute("id", "toolbar");
    
		var img = document.createElement("img");
		img.setAttribute("src", "images/new.jpg");
		img.setAttribute("alt", "New Sheet");
		img.setAttribute("title", img.getAttribute("alt"));
		img.setAttribute("border", "0");
		img.onmouseover = function () { this.setAttribute("src", "images/new_hover.jpg"); };
		img.onmouseout = function () { this.setAttribute("src", "images/new.jpg"); };

		var link = document.createElement("a");
		link.setAttribute("href", "javascript:void(0);");
		link.onclick = function () {
			_this.addSheet();
		};
		link.appendChild(img);
		div.appendChild(link);

		var img = document.createElement("img");
		img.setAttribute("src", "images/save.jpg");
		img.setAttribute("alt", "Save");
		img.setAttribute("title", img.getAttribute("alt"));
		img.setAttribute("border", "0");
		img.onmouseover = function () { this.setAttribute("src", "images/save_hover.jpg"); };
		img.onmouseout = function () { this.setAttribute("src", "images/save.jpg"); };

		var link = document.createElement("a");
		link.setAttribute("href", "javascript:void(0);");
		link.onclick = function () { _this.saveData(); };
		link.appendChild(img);
		div.appendChild(link);

		var img = document.createElement("img");
		img.setAttribute("src", "images/print.jpg");
		img.setAttribute("alt", "Print");
		img.setAttribute("title", img.getAttribute("alt"));
		img.setAttribute("border", "0");
		img.onmouseover = function () { this.setAttribute("src", "images/print_hover.jpg"); };
		img.onmouseout = function () { this.setAttribute("src", "images/print.jpg"); };

		var link = document.createElement("a");
		link.setAttribute("href", "javascript:void(0);");
		link.onclick = function () { alert("Print"); };
		link.appendChild(img);
		div.appendChild(link);

		var img = document.createElement("img");
		img.setAttribute("src", "images/colsep.jpg");
		img.setAttribute("border", "0");
		div.appendChild(img);

		var img = document.createElement("img");
		img.setAttribute("src", "images/insert_column.jpg");
		img.setAttribute("alt", "Insert Column");
		img.setAttribute("title", img.getAttribute("alt"));
		img.setAttribute("border", "0");
		img.onmouseover = function () { this.setAttribute("src", "images/insert_column_hover.jpg"); };
		img.onmouseout = function () { this.setAttribute("src", "images/insert_column.jpg"); };

		var link = document.createElement("a");
		link.setAttribute("href", "javascript:void(0);");
		link.onclick = function () {
			_this.addColumn();
		};
		link.appendChild(img);
		div.appendChild(link);

		var img = document.createElement("img");
		img.setAttribute("src", "images/insert_row.jpg");
		img.setAttribute("alt", "Insert Row");
		img.setAttribute("title", img.getAttribute("alt"));
		img.setAttribute("border", "0");
		img.onmouseover = function () { this.setAttribute("src", "images/insert_row_hover.jpg"); };
		img.onmouseout = function () { this.setAttribute("src", "images/insert_row.jpg"); };

		var link = document.createElement("a");
		link.setAttribute("href", "javascript:void(0);");
		link.onclick = function () {
			_this.addRow();
		};
		link.appendChild(img);
		div.appendChild(link);

		var img = document.createElement("img");
		img.setAttribute("src", "images/delete.jpg");
		img.setAttribute("alt", "Delete");
		img.setAttribute("title", img.getAttribute("alt"));
		img.setAttribute("border", "0");
		img.onmouseover = function () { this.setAttribute("src", "images/delete_hover.jpg"); };
		img.onmouseout = function () { this.setAttribute("src", "images/delete.jpg"); };

		var link = document.createElement("a");
		link.setAttribute("href", "javascript:void(0);");
		link.onclick = function () {
			if (_this.selectedRow == -1 && _this.selectedColumn == -1) {
				alert("You must select a row or column to delete first");
			}
			else {
				if (_this.selectedColumn != -1)
					_this.delColumn(_this.selectedColumn);
				else
					_this.delRow(_this.selectedRow);
			}
		};
		link.appendChild(img);
		div.appendChild(link);

		var img = document.createElement("img");
		img.setAttribute("src", "images/colsep.jpg");
		img.setAttribute("border", "0");
		div.appendChild(img);

		var _this = this;
		var newLink = document.createElement("a");
		newLink.setAttribute("href", "javascript:void(0);");
		newLink.setAttribute("id", this.selectedSheet);
		newLink.onclick = function () {
			var cfm = confirm("Are you sure you want to delete the worksheet titled \"" + _this.rows[_this.selectedSheet][0] + "\"?");

			if (cfm) {
				var newIndex = 0;
				if ((_this.selectedSheet - 1 < 0) && (_this.selectedSheet + 1 < _this.rows.length) && (_this.selectedSheet != 0))
					newIndex = _this.selectedSheet + 1;
				else if (_this.selectedSheet - 1 >= 0)
					newIndex = _this.selectedSheet - 1;
				_this.selectSheet(newIndex);

				_this.delSheet(this.getAttribute("id"));
				_this.draw();
			}
		};

		var img = document.createElement("img");
		img.setAttribute("src", "images/delete_sheet.jpg");
		img.setAttribute("alt", "Delete Active Worksheet");
		img.setAttribute("title", img.getAttribute("alt"));
		img.setAttribute("border", "0");
		img.onmouseover = function () { this.setAttribute("src", "images/delete_sheet_hover.jpg"); };
		img.onmouseout = function () { this.setAttribute("src", "images/delete_sheet.jpg"); };
		newLink.appendChild(img);
		div.appendChild(newLink);

		var _this = this;
		var newLink = document.createElement("a");
		newLink.setAttribute("href", "javascript:void(0);");
		newLink.setAttribute("id", this.selectedSheet);
		newLink.onclick = function () {
			var cfm = prompt("Please enter a new name for the worksheet \"" + _this.rows[_this.selectedSheet][0] + "\"", _this.rows[_this.selectedSheet][0]);
			if (cfm && cfm.length > 0) {
				_this.rows[_this.selectedSheet][0] = cfm;
				_this.draw();
			}
			else if (cfm) {
				alert("Invalid worksheet name");
			}
		};

		var img = document.createElement("img");
		img.setAttribute("src", "images/rename_sheet.jpg");
		img.setAttribute("alt", "Rename Active Worksheet");
		img.setAttribute("title", img.getAttribute("alt"));
		img.setAttribute("border", "0");
		img.onmouseover = function () { this.setAttribute("src", "images/rename_sheet_hover.jpg"); };
		img.onmouseout = function () { this.setAttribute("src", "images/rename_sheet.jpg"); };
		newLink.appendChild(img);
		div.appendChild(newLink);

		var img = document.createElement("img");
		img.setAttribute("src", "images/colsep.jpg");
		img.setAttribute("border", "0");
		div.appendChild(img);

		var img = document.createElement("img");
		img.setAttribute("src", "images/sort_asc.jpg");
		img.setAttribute("alt", "Sort Ascending");
		img.setAttribute("title", img.getAttribute("alt"));
		img.setAttribute("border", "0");
		img.onmouseover = function () { this.setAttribute("src", "images/sort_asc_hover.jpg"); };
		img.onmouseout = function () { this.setAttribute("src", "images/sort_asc.jpg"); };

		var link = document.createElement("a");
		link.setAttribute("href", "javascript:void(0);");
		link.onclick = function () {
			if (_this.selectedColumn == -1)
				alert("You must select a column to sort first");
			else
				alert("Sort Ascending");
		};
		link.appendChild(img);
		div.appendChild(link);

		var img = document.createElement("img");
		img.setAttribute("src", "images/sort_desc.jpg");
		img.setAttribute("alt", "Sort Descending");
		img.setAttribute("title", img.getAttribute("alt"));
		img.setAttribute("border", "0");
		img.onmouseover = function () { this.setAttribute("src", "images/sort_desc_hover.jpg"); };
		img.onmouseout = function () { this.setAttribute("src", "images/sort_desc.jpg"); };

		var link = document.createElement("a");
		link.setAttribute("href", "javascript:void(0);");
		link.onclick = function () {
			if (_this.selectedColumn == -1)
				alert("You must select a column to sort first");
			else
				alert("Sort Descending");
		};
		link.appendChild(img);
		div.appendChild(link);

		// Function reference DIV - Debugging only. This will be
		// added as a button in the toolbar

		var img = document.createElement("img");
		img.setAttribute("src", "images/colsep.jpg");
		img.setAttribute("border", "0");
		div.appendChild(img);

		var span = document.createElement("span");
		var txt = document.createTextNode("Function lookup: ");
		span.setAttribute("style", "vertical-align: middle; display: inline-block; margin-top: -15px");
		span.appendChild(txt);
		div.appendChild(span);

		var _this = this;
		var input = document.createElement("input");
		input.setAttribute("type", "text");
		input.setAttribute("id", "BNGrid_FuncRefInput");
		input.setAttribute("style", "vertical-align: top;");

		input.onkeypress = function (e) {
			if (e.keyCode != 13)
				return;

			var txt = document.getElementById("BNGrid_FuncRefInput");
			var val = txt.value.replace(/^=/, "").toUpperCase();

			var i;
			for (i = 0; i < _this.supportedExcelFunctions.length; i++) {
				if (val == _this.supportedExcelFunctions[i][0]) {
					var alt = _this.supportedExcelFunctions[i][0];

					if (_this.supportedExcelFunctions[i].length > 3)
						alt += _this.supportedExcelFunctions[i][2];
					else
						alt += "()";

					alt += "\n";
					alt += _this.supportedExcelFunctions[i][1];
					alert(alt);
					return;
				}
			}

			alert("Function " + val + " was not found.");
		};
		div.appendChild(input);

		input = document.createElement("input");
		input.setAttribute("type", "button");
		input.setAttribute("value", "Go");
		input.setAttribute("style", "vertical-align: top;");
		input.onclick = function () {
			var txt = document.getElementById("BNGrid_FuncRefInput");

			txt.value.replace(/^=/, "");
			var val = txt.value.replace(/^=/, "").toUpperCase();

			var i;
			for (i = 0; i < _this.supportedExcelFunctions.length; i++) {
				if (val == _this.supportedExcelFunctions[i][0]) {
					var alt = _this.supportedExcelFunctions[i][0];

					if (_this.supportedExcelFunctions[i].length > 3)
						alt += _this.supportedExcelFunctions[i][2];
					else
						alt += "()";

					alt += "\n";
					alt += _this.supportedExcelFunctions[i][1];
					alert(alt);
					return;
				}
			}

			alert("Function " + val + " was not found.");
		};
	div.appendChild(input);
	ct.appendChild(div);

	// Formula bar
	var div = document.createElement("div");
	div.setAttribute (
		"style",
		"background: url(images/fxbg.jpg); " +
		"display: block; " +
		"height: 24px; " +
		"padding-top: 2px;"
	);

	var tab = document.createElement("table");
	tab.setAttribute("cellpadding", "0");
	tab.setAttribute("cellspacing", "0");
	tab.setAttribute("border", "0");
	tab.setAttribute("style", "width: 100%;");

	var tr = document.createElement("tr");
	tab.appendChild(tr);

	var formEdit = document.createElement("input");
	formEdit.setAttribute("type", "text");
	formEdit.setAttribute("style", "border: 1px solid #000000; width: 100%; height: 22px; margin-top: -4px;");
	formEdit.setAttribute("id", "formulaEditBox");
    
	var _this = this;
	formEdit.onchange = function () {
		if (_this.selectedCell !== undefined) {
			var tmp = _this.selectedCell.split(";");

			tmp[0] = tmp[0].replace("row", "");
			tmp[1] = tmp[1].replace("col", "");

			tmp[0] = parseInt(tmp[0]) - 1;
			tmp[1] = parseInt(tmp[1]) - 1;

			if (this.value.match(/^=/)) {
				_this.setFormula(tmp[0], tmp[1], this.value);
			}
			else {
				_this.setValue(tmp[0], tmp[1], this.value);
				_this.setFormula(tmp[0], tmp[1], "");
			}

			_this.draw();

			_this.highlightHeader(_this.selectedCell);
			_this.highlightCell(document.getElementById(_this.selectedCell), _this);

			var formula = _this.getFormula(tmp[0], tmp[1]);

			if (formula.length == 0)
				formula = _this.getValue(tmp[0], tmp[1]);

			document.getElementById("formulaEditBox").setAttribute("value", formula);
			document.getElementById("formulaEditBox").focus();
		}
	}

	formEdit.onkeypress = function (e) {
		if (e.keyCode == 13) {
			this.onchange();
			return;
		}
	}

	var td = document.createElement("td");
	td.setAttribute("style", "padding-right: 4px; width: 14px; text-align: right;");

	var link = document.createElement("a");
	link.setAttribute (
		"style",
		"text-decoration: none; " +
		"padding-left: 2px; " +
		"margin-right: -5px;"
	);
	link.setAttribute("href", "javascript:void(0);");
	link.onclick = function () {
		var id = document.getElementById("functionRefAll");

		if (id === undefined)
			return;

		if (id.style.display == "none") {
			id.style.top = this.clientY;
			id.style.left = this.clientX;

			id.style.display = "block";

			var img = document.getElementById("imgFx");
			img.setAttribute("src", "images/fx_on.jpg");
		}
		else {
			id.style.display = "none";
			var img = document.getElementById("imgFx");
			img.setAttribute("src", "images/fx.jpg");
		}
	};

	var img = document.createElement("img");
	img.setAttribute("border", "0");
	img.setAttribute("src", "images/fx.jpg");
	img.setAttribute("style", "vertical-align: top; display: inline-block; margin-top: -5px; margin-left: -5px;");
	img.setAttribute("id", "imgFx");
	link.appendChild(img);
	td.appendChild(link);

	var newDiv = document.createElement("div");
	newDiv.setAttribute("style",
		"position: absolute; " +
		"border: 1px solid #000000; " +
		"display: none; " +
		"background-color: #FFFFFF; " +
		"overflow: auto; " +
		"max-height: 200px; " +
		"max-width: 400px;" +
		"text-align: left;"
	);

	newDiv.setAttribute("id", "functionRefAll");

	var i;
	var tbl = document.createElement("table");
	for (i = 0; i < this.supportedExcelFunctions.length; i++) {
		var funcname = this.supportedExcelFunctions[i][0];
      
		if (this.supportedExcelFunctions[i].length == 4)
			funcname += this.supportedExcelFunctions[i][2];
		else
			funcname += "()";

		var tr1 = document.createElement("tr");
		var td1 = document.createElement("td");

		td1.innerHTML = funcname + "<br><em>" + this.supportedExcelFunctions[i][1] + "</em>";

		tr1.appendChild(td1);
		tbl.appendChild(tr1);

		if (i < (this.supportedExcelFunctions.length - 1)) {
			var tr1 = document.createElement("tr");
			var td1 = document.createElement("td");
			td1.setAttribute("style", "border-bottom: 1px solid #CCCCCC;");
			tr1.appendChild(td1);
			tbl.appendChild(tr1);
		}
	}

	newDiv.appendChild(tbl);
	td.appendChild(newDiv)
	tr.appendChild(td);

	td = document.createElement("td");
	tr.appendChild(td);
	td.appendChild(formEdit);
	div.appendChild(tab);
	ct.appendChild(div);

	div = document.createElement("div");
	div.setAttribute("style", "max-height: 200px; overflow: auto;");
	div.setAttribute("id", "BNGridDataArea");
	div.appendChild(table);
	ct.appendChild(div);

		// Now generate the links for each sheet in this workbook
		var i;
		for (i = 0; i < this.rows.length; i++) {
			var div = document.createElement("div");

			var style = "";
			if (i == this.selectedSheet) {
				style +=	"background-color: #e2ebf2; " +
							"background: url(images/background.jpg); " +
							"padding-top: 4px; padding-bottom: 4px; ";
			}
	
			style +=	"min-width: 50px; "  +
						"display: box; " +
						"border: 1px solid #000000; " +
						"float: left; " +
						"padding: 2px; " +
						"background-color: #e2ebf2; " +
						"border-top: 0px; ";
			if ((i != 0) && ((i + 1) == this.rows.length))
				style += "margin-left: -1px; ";
			else if ((i != 0) && ((i + 1) != this.rows.length))
				style += "margin-left: -1px; ";
			div.setAttribute("style", style);

			nTxt = document.createTextNode(" ");
			div.appendChild(nTxt);

			var link = document.createElement("a");
			link.setAttribute("href", "javascript:void(0);");
			link.setAttribute("id", i);

			if (i == this.selectedSheet) {
				link.setAttribute("style", "display: block; text-decoration: none; color: #000000; font-weight: bold;");
				link.ondblclick = function () {
					var cfm = prompt(
						"Please enter a new name for the worksheet \"" +
						_this.rows[_this.selectedSheet][0] + "\"",
						_this.rows[_this.selectedSheet][0]
					);

					if (cfm && cfm.length > 0) {
						_this.rows[_this.selectedSheet][0] = cfm;
						_this.draw();
					}
					else if (cfm) {
						alert("Invalid worksheet name");
					}
				};
			}
			else {
				link.setAttribute("style", "display: block; text-decoration: none; color: #000000;");
				link.onclick = function () { _this.selectSheet(this.getAttribute("id")); };
			}

			var txt = document.createTextNode(this.rows[i][0]);
			link.appendChild(txt);
			div.appendChild(link);

			ct.appendChild(div);
		}

		var div = document.createElement("div");
		div.setAttribute("style", "clear: both;");

		var txt = document.createTextNode("BNGrid");
		var link = document.createElement("a");
		link.setAttribute("href", "http://www.blendernet.com");
		link.appendChild(txt);

		div.innerHTML = "<br><hr>";

		txt = document.createTextNode("Generated by ");

		div.appendChild(txt);
		div.appendChild(link);

		div.innerHTML += ". &copy; 2009 BlenderNet, LLC. All rights reserved.";

		ct.appendChild(div);
	}
}
