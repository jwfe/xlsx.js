const xlsxStyle = require('xlsx-style');


class XLSX {
	constructor(otions = {}) {
		this.xsModel = {
			workbook: []
		};

		this.merges = [];

		this.workbook = {
			SheetNames: [options.xlsxName],
			Sheets: {}
		};

		this.workbook.Sheets[options.xlsxName] = this.xsModel.workbook;

		this.style = {
			fontSize: options.style.fontSize || 9,
			fontFamily: options.style.fontFamily || 'Microsoft YaHei',
			backgroundColor: options.style.backgroundColor || 'FFF2F2F2',
			td: {
				alignment: {
					horizontal: this.tdHorizontal || 'center',
					vertical: this.tdVertical || 'center',
					wrapText: true
				},
				border: {
					left: {
						style: this.border.style || 'thin',
						color: {
							rgb: this.border.color || 'FFBEBEBE'
						}
					},
					right: {
						style: this.border.style || 'thin',
						color: {
							rgb: this.border.color || 'FFBEBEBE'
						}
					},
					top: {
						style: this.border.style || 'thin',
						color: {
							rgb: this.border.color || 'FFBEBEBE'
						}
					},
					bottom: {
						style: this.border.style || 'thin',
						color: {
							rgb: this.border.color || 'FFBEBEBE'
						}
					}
				}
			}
		}
	}

	style(cell, config){
		if (config.style) {
			for (let key in config.style) {
				cell.s[key] = config.style[key];
			}
		}
	}

	addMerge(range, config = {}){
		this.merges.push(range);
		//todo: 添加空空列， 用于样式
		for(let c = range.s.c; c <= range.e.c; c++ ) {
			for(let r = range.s.r; r <= range.e.r; r++ ) {
				if(!(c === range.s.c && r === range.s.r)) {
					config.col = c;
					config.row = r;
					this.addCell('', config)
				}
			}
		}
	}

	addCell(val, config) {

		let type = typeof val;
		type = type.substr(0, 1);
		if (val == null) {
			val = ''
		}
		val = val.toString().replace(/<[^>]+>/img, '').replace('&nbsp;', '');
		if (type == 'n') {
			type = 's';
		}
		let cell = {
			v: val,
			t: type,
			s: {
				font: {sz: this.style.fontSize, name: this.style.fontFamily},
				alignment: this.style.td.alignment
			}
		};
		if (!config.noBorder) {
			cell.s.border = this.style.td.border;
		}

		this.style(cell, config);

		if (config.merges) {
			const merges = config.merges;
			const _conf = {};
			for (let k in config) {
				if (!(k == 'merges' || k == 'col' || k == 'row')) {
					_conf[k] = config[k];
				}
			}
			this.addMerge(merges, _conf);
		}
		const ref = XLSX.utils.encode_cell({c: config.col, r: config.row});
		this.xsModel.workbook[ref] = cell;
	}

	buffer() {
		const config = {
			ext: 'xlsx',
			bookSST: false,
			encoding: 'binary',
			showGridLines: false
		};

		return new Buffer(xlsxStyle.write(this.workbook, {
			bookType: config.ext,
			bookSST: false,
			type: config.encoding,
			showGridLines: false
		}), config.encoding)
	}

}

module.exports = XLSX;
