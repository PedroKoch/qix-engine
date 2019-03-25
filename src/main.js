const enigma = require('enigma.js');
const WebSocket = require('ws');
const path = require('path');
const fs = require('fs');
const schema = require('enigma.js/schemas/12.20.0.json');
const xlsx = require('xlsx');
const readline = require('readline');

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
  prompt: 'NS> '
});

function dynamicSort(property) {
    var sortOrder = 1;

    if(property[0] === "-") {
        sortOrder = -1;
        property = property.substr(1);
    }

    return function (a,b) {
        if(sortOrder == -1){
            return b[property].localeCompare(a[property]);
        }else{
            return a[property].localeCompare(b[property]);
        }
    }
}

function addDim(ID, Name) {
  var obj = {
    "qLibraryId": "" + ID,
    "qDef": {
      "qGrouping": "N",
      "qFieldDefs": [],
      "qFieldLabels": [],
      "qSortCriterias": [
        {
          "qSortByNumeric": 1,
          "qSortByAscii": 1,
          "qSortByLoadOrder": 1,
          "qExpression": {}
        }
      ],
      "qNumberPresentations": [],
      "qActiveField": 0,
      "autoSort": true,
      "cId": "" + ID.substring(2,6) + Name.substring(1,4),
      "othersLabel": "Outros",
      "textAlign": {
        "auto": true,
        "align": "left"
      },
      "representation": {
        "type": "text",
        "urlLabel": ""
      }
    },
    "qOtherTotalSpec": {
      "qOtherMode": "OTHER_OFF",
      "qOtherCounted": {
        "qv": "10"
      },
      "qOtherLimit": {
        "qv": "0"
      },
      "qOtherLimitMode": "OTHER_GE_LIMIT",
      "qForceBadValueKeeping": true,
      "qApplyEvenWhenPossiblyWrongResult": true,
      "qOtherSortMode": "OTHER_SORT_DESCENDING",
      "qTotalMode": "TOTAL_OFF",
      "qReferencedExpression": {}
    },
    "qOtherLabel": {
      "qv": "Outros"
    },
    "qTotalLabel": {},
    "qCalcCond": {},
    "qAttributeExpressions": [
      {
        "qExpression": "",
        "id": "cellBackgroundColor"
      },
      {
        "qExpression": "",
        "id": "cellForegroundColor"
      }
    ],
    "qAttributeDimensions": [],
    "qCalcCondition": {
      "qCond": {
        "qv": "if(Count({<[_Dim] *= {'" + Name + "'}>} [_Dim]) > 0 and GetSelectedCount([_Dim]) > 0, 1, 0)",
      },
      "qMsg": {}
    }
  };
  return obj;
}

function addInd(ID, Name) {
  var obj = {
        "qLibraryId": "" + ID,
        "qDef": {
          "qTags": [],
          "qGrouping": "N",
          "qNumFormat": {
            "qType": "U",
            "qnDec": 10,
            "qUseThou": 0
          },
          "qAggrFunc": "Expr",
          "qAccumulate": 0,
          "qActiveExpression": 0,
          "qExpressions": [],
          "autoSort": true,
          "cId": "" + ID.substring(2,6) + Name.substring(1,4),
          "numFormatFromTemplate": true,
          "textAlign": {
            "auto": true,
            "align": "left"
          }
        },
        "qSortBy": {
          "qSortByNumeric": -1,
          "qSortByLoadOrder": 1,
          "qExpression": {}
        },
        "qAttributeExpressions": [
          {
            "qExpression": "",
            "id": "cellBackgroundColor"
          },
          {
            "qExpression": "",
            "id": "cellForegroundColor"
          }
        ],
        "qAttributeDimensions": [],
        "qCalcCond": {},
        "qCalcCondition": {
          "qCond": {
            "qv": "if(Count({<[_Ind] *= {'" + Name + "'}>} [_Ind]) > 0 and GetSelectedCount([_Ind]) > 0, 1, 0)"
          },
          "qMsg": {}
        }
      };
  return obj;
}

var menuHandler = function(input) {
  switch(input) {
    case 'm': menu(); break;
    case '1': menuExcel(); break;
    case '2': menuQS(); break;
    case '':
      rl.prompt();
      break;
    case '0':
      console.log('\n Have a nice day!!');
      process.exit(0);
      break;
    default:
      console.log("'" + input + "' is not a known command.");
      rl.prompt();
  }
};

var excel = {
  "init": 0,
  reset: function() {
    for(i in this) {
      if(typeof(this[i]) != "function") {
        console.log("Deleting " + i);
        delete this[i];
      }
    }
    this.init = 0;
  },
  open: function(file) {
    this.file = file;
    this.wb = xlsx.readFile(file);
    if(!(this.wb == undefined || this.wb == null)) {
      this.init = 1;
      console.log('Opened file: ' + file);
    }
  },
  listSheets: function() {
    if(this.init > 0) {
      var qt = 0;
      console.log('Opened file: ' + this.file);
      console.log('\nChoose one of the following sheets: ');
      for(s in this.wb.SheetNames) {
        qt++;
        console.log('   ' + (+s+1) + ' - ' + this.wb.SheetNames[s]);
      }
      this.nsheets = qt;
      console.log('\n');
    }
  },
  setSheet: function(idx) {
    this.currSheet = this.wb.SheetNames[idx];
    this.ws = this.wb.Sheets[this.currSheet];
    if(this.currSheet == undefined) {
      console.log('Sheet does not exist.');
      return 0;
    } else {
      console.log('Sheet ' + this.currSheet + ' opened.\n');
    }
    this.sheet = xlsx.utils.sheet_to_json(this.ws);
    this.sheet.cols = {
      "title": Object.keys(this.sheet[0])
    };
    this.data = {
      "meta": {
        "colnames": [],
        "ncols": 0,
        "nrows": 0
      },
      element: function (x, y) {
        return this[this.meta.colnames[x]][y];
      }
    };
    this.__update();
    if(!(this.ws == undefined || this.ws == null)) {
      this.init = 2;
    }
  },
  __update: function() {
    var qt = 0;
    for(i in this.sheet.cols.title) {
      this.data[this.sheet.cols.title[i]] = [];
      this.data.meta.colnames.push(this.sheet.cols.title[i]);
      qt += 1;
    };
    this.data.meta.ncols = qt;
    qt = 0;
    for(i in this.sheet) { //lines
      for(prop in this.sheet[i]) { //each element of one line
        if(this.data.meta.colnames.indexOf(prop) > -1) {
          this.data[prop].push(this.sheet[i][prop]);
        }
      };
      qt += 1;
    };
    this.data.meta.nrows = qt>0 ? qt-1 : 0;
  },
  preview: function(opt = 10) {
    var max = [], rows = Math.min(Number(isNaN(opt) || !(opt>0) ? 10 : opt), this.data.meta.nrows);
    for(i = 0; i < this.data.meta.ncols; i++) {
      max.push(0);
    }
    for(i = 0; i < rows; i++) {
      for(j = 0; j < this.data.meta.ncols; j++) {
        var aux = String(this.data[this.data.meta.colnames[j]][i]).length;
        if(max[j] < aux) {
          max[j] = aux;
        }
      }
    }
    var acum = 0;
    for(i in max) {
      max[i] = Math.floor(Math.min(Math.max(1.5 * max[i], 2 + this.data.meta.colnames[i].length) , (process.stdout.columns - acum) / (max.length-i)));
      acum += max[i];
    }
    this.data.meta.pcolw = max;
    console.log(' Preview of ' + this.currSheet + ' (' + this.file + ')\n');
    var cab = ' ';
    for(j = 0; j < this.data.meta.ncols; j++) {
      cab += this.data.meta.colnames[j];
      if(j < this.data.meta.ncols - 1) {
        cab += ' '.repeat(this.data.meta.pcolw[j] - this.data.meta.colnames[j].length);
      }
    }
    console.log(cab);
    for(i = 0; i < rows; i++) {
      var linha = '';
      for(j = 0; j < this.data.meta.ncols; j++) {
        var elem = this.data.element(j,i);
        elem = (typeof(elem) === undefined ? '' : String(elem));
        elem = (elem.length > this.data.meta.pcolw[j] - 1 ? elem.substring(0,this.data.meta.pcolw[j]-5) + '...' : elem);
        linha += elem;
        if(j < this.data.meta.ncols - 1) {
          linha += ' '.repeat(this.data.meta.pcolw[j] - elem.length);
        }
      }
      console.log(linha);
    }
  }
};

var QIX = {
  "init": 0,
  ini: function() {
    let rawconfig = fs.readFileSync('./config.json'); 
	let config = JSON.parse(rawconfig);
    this.engineHost = config.engineHost;
    this.enginePort = config.enginePort;
    this.appId = config.appId;
    this.userDirectory = config.userDirectory;
    this.userId = config.userId;
    this.certificatesPath = config.certificatesPath;
	this.server = config.server;
	if(this.server) {
	  readCert = filename => fs.readFileSync(path.resolve(__dirname, this.certificatesPath, filename));
	  this.session = enigma.create({
		schema,
		url: `wss://${this.engineHost}:${this.enginePort}/app/${this.appId}`,
		createSocket: url => new WebSocket(url, {
		  ca: [readCert('root.pem')],
		  key: readCert('client_key.pem'),
		  cert: readCert('client.pem'),
		  headers: {
		    'X-Qlik-User': `UserDirectory=${encodeURIComponent(this.userDirectory)}; UserId=${encodeURIComponent(this.userId)}`,
		  },
	    })
	  })
	} else {
	  this.session = enigma.create({
		schema,
		url: 'ws://localhost:4848/app/engineData',
		createSocket: url => new WebSocket(url)
	  })
	}
    
    process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";
    this.received = [];
    this.sent = [];
    this.session.on('traffic:sent', data => {this.sent.push(data); /*console.log('sent:', data);*/});
    this.session.on('traffic:received', data => {this.received.push(data); /*console.log('received:', this.received[this.received.length-1]);*/});
    this.init = 1;
  },
  try: function(AppID) {
    if(this.init == 1) {
      this.session.open().then((global) => {
        console.log('\nSession opened.');
        global.openDoc(AppID).then((doc) => {
          doc.getAppLayout()
          .then((layout) => {
            console.log('Connected to app ' + layout.qTitle + '.\n');
            this.AppID = AppID;
            this.AppName = layout.qTitle;
          });
        })
        .catch(() => {console.log('It is not possible to connect to ' + AppID + '.\n')});
      });
    }
  },
  on: function(control, AppID = this.AppID) {
    if(this.init == 1) {
      switch (control) {
        case 'List Measures':
          this.session.open().then((global) => {
            global.openDoc(AppID).then((doc) => {
              var param = {
          			"qInfo": {
          				"qType": "MeasureList"
          			},
          			"qMeasureListDef": {
          				"qType": "measure",
          				"qData": {
          					"title": "/title",
          					"tags": "/tags"
          				}
          			}
          		};
              doc.createSessionObject(param)
              .then((mList) => {
                mList.getLayout()
                .then((layout) => {
                  QIX.MeasureList = [];
                  console.log('\n  **  ' + (layout.qMeasureList.qItems.length > 0 ? layout.qMeasureList.qItems.length : 0) + ' Measures Found  **  \n');
                  for(i in layout.qMeasureList.qItems) {
                    QIX.MeasureList.push({ID: layout.qMeasureList.qItems[i].qInfo.qId, Título: layout.qMeasureList.qItems[i].qMeta.title});
                    console.log(QIX.MeasureList[i].Título + ' (' + QIX.MeasureList[i].ID + ')');
                  }
                  console.log(' ');
                });
              });
            });
          });
          break;
        case 'List Dimensions':
          this.session.open().then((global) => {
            global.openDoc(AppID).then((doc) => {
              var param = {
          			"qInfo": {
          				"qType": "DimensionList"
          			},
          			"qDimensionListDef": {
          				"qType": "dimension",
          				"qData": {
          					"title": "/title",
          					"tags": "/tags",
          					"grouping": "/qDim/qGrouping",
          					"info": "/qDimInfos"
          				}
          			}
          		};
              doc.createSessionObject(param)
              .then((dList) => {
                dList.getLayout()
                .then((layout) => {
                  QIX.DimensionList = [];
                  console.log('\n  **  ' + (layout.qDimensionList.qItems.length > 0 ? layout.qDimensionList.qItems.length : 0) + ' Dimensions Found  **  \n');
                  for(i in layout.qDimensionList.qItems) {
                    QIX.DimensionList.push({ID: layout.qDimensionList.qItems[i].qInfo.qId, Título: layout.qDimensionList.qItems[i].qMeta.title});
                    console.log(QIX.DimensionList[i].Título + ' (' + QIX.DimensionList[i].ID + ')');
                  }
                  console.log(' ');
                });
              });
            });
          });
          break;
        case 'List Variables':
          this.session.open().then((global) => {
            global.openDoc(AppID).then((doc) => {
              var param = {
        				"qInfo": {
        					"qType": "VariableList"
        				},
        				"qVariableListDef": {
        					"qType": "variable",
        					"qShowReserved": true,
        					"qShowConfig": true,
        					"qData": {
        						"tags": "/tags"
        					}
        				}
        			};
              doc.createSessionObject(param)
              .then((vList) => {
                vList.getLayout()
                .then((layout) => {
                  QIX.VariableList = [];
                  console.log('\n  **  ' + (layout.qVariableList.qItems.length > 0 ? layout.qVariableList.qItems.length : 0) + ' Variables Found  **  \n');
                  var j = 0;
                  for(i in layout.qVariableList.qItems) {
                    if(!("qIsReserved" in layout.qVariableList.qItems[i]) && !("qIsConfig" in layout.qVariableList.qItems[i]) ) {
                      QIX.VariableList.push({ID: layout.qVariableList.qItems[i].qInfo.qId, Título: layout.qVariableList.qItems[i].qName});
                      console.log(QIX.VariableList[j].Título + ' (' + QIX.VariableList[j].ID + ')');
                      j++;
                    }
                  }
                  console.log(' ** Showing only variables created by the user (' + j + ').\n');
                });
              });
            });
          });
          break;
        case 'Delete All Measures':
          this.session.open().then((global) => {
            global.openDoc(AppID).then((doc) => {
              var param = {
                "qInfo": {
                  "qType": "MeasureList"
                },
                "qMeasureListDef": {
                  "qType": "measure",
                  "qData": {
                    "title": "/title",
                    "tags": "/tags"
                  }
                }
              };
              doc.createSessionObject(param)
              .then((mList) => {
                mList.getLayout()
                .then((layout) => {
                  QIX.MeasureList = [];
                  for(i in layout.qMeasureList.qItems) {
                    QIX.MeasureList.push({ID: layout.qMeasureList.qItems[i].qInfo.qId, Título: layout.qMeasureList.qItems[i].qMeta.title});
                    doc.destroyMeasure({'qId': QIX.MeasureList[i].ID});
                  }
                  console.log('Done.');
                });
              });
            });
          });
          break;
        case 'Delete All Dimensions':
          this.session.open().then((global) => {
            global.openDoc(AppID).then((doc) => {
              var param = {
          			"qInfo": {
          				"qType": "DimensionList"
          			},
          			"qDimensionListDef": {
          				"qType": "dimension",
          				"qData": {
          					"title": "/title",
          					"tags": "/tags",
          					"grouping": "/qDim/qGrouping",
          					"info": "/qDimInfos"
          				}
          			}
          		};
              doc.createSessionObject(param)
              .then((dList) => {
                dList.getLayout()
                .then((layout) => {
                  QIX.DimensionList = [];
                  for(i in layout.qDimensionList.qItems) {
                    QIX.DimensionList.push({ID: layout.qDimensionList.qItems[i].qInfo.qId, Título: layout.qDimensionList.qItems[i].qMeta.title});
                    doc.destroyDimension({'qId': QIX.DimensionList[i].ID});
                  }
                  console.log('Done.');
                });
              });
            });
          });
          break;
        case 'Delete All Variables':
          this.session.open().then((global) => {
            global.openDoc(AppID).then((doc) => {
              var param = {
        				"qInfo": {
        					"qType": "VariableList"
        				},
        				"qVariableListDef": {
        					"qType": "variable",
        					"qShowReserved": true,
        					"qShowConfig": true,
        					"qData": {
        						"tags": "/tags"
        					}
        				}
        			};
              doc.createSessionObject(param)
              .then((vList) => {
                vList.getLayout()
                .then((layout) => {
                  QIX.VariableList = [];
                  for(i in layout.qVariableList.qItems) {
                    if(!("qIsReserved" in layout.qVariableList.qItems[i])) {
                      QIX.VariableList.push({ID: layout.qVariableList.qItems[i].qInfo.qId, Título: layout.qVariableList.qItems[i].qName});
                      doc.destroyVariableById({'qId': QIX.VariableList[i].ID});
                    }
                  }
                  console.log('Done.');
                });
              });
            });
          });
          break;
        case 'Load Measures':
          this.session.open().then((global) => {
            global.openDoc(AppID).then((doc) => {
              console.log('Loading ' + excel.data.meta.nrows + ' lines.');
              for(i = 0; i < excel.data.meta.nrows; i++) {
                var param = {
              		"qProp": {
              			"qInfo": {
              				"qId": "",
              				"qType": "measure"
              			},
                    "qMeasure": {
              				"qLabel": "" + excel.data.element(1,i), //measure name, the one that shows in Edit Mode config panel
              				"qDef": "" + excel.data.element(3,i), //measure definition
              				"qGrouping": 0,
              				"qExpressions": [
              					""
              				],
              				"qActiveExpression": 0,
              				"qLabelExpression": "='" + excel.data.element(2,i) + "'" //measure label
              			},
              			"qMetaDef": {
                      "title": "" + excel.data.element(1,i), //measure name
                      "description": "" + excel.data.element(4,i), //measure description
                      "tags": [
                        "ID" + excel.data.element(0,i),
                        "engine_api",
                        "created_by_nodejs"
                      ]
                    }
              		}
              	};
                doc.createMeasure(param);
              }
            });
          });
          break;
        case 'Load Dimensions':
          this.session.open().then((global) => {
            global.openDoc(AppID).then((doc) => {
              console.log('Loading ' + excel.data.meta.nrows + ' lines.');
              for(i = 0; i < excel.data.meta.nrows; i++) {
                var param = {
              		"qProp": {
              			"qInfo": {
              				"qId": "",
              				"qType": "dimension"
              			},
              			"qDim": {
              				"qGrouping": 0,
              				"qFieldDefs": [
              					"["+ excel.data.element(2,i) +"]" //field definition
              				],
              				"qFieldLabels": [
              					""
              				],
              				"qLabelExpression": "='" + excel.data.element(4,i) + "'", //dimension label
                      "title": "" + excel.data.element(1,i) //dimension name, the one that shows in Edit Mode config panel
              			},
              			"qMetaDef": {
                      "title": "" + excel.data.element(1,i), //dimension name
                      "description": "Field in table " + excel.data.element(3,i),
                      "tags": [
                        "" + excel.data.element(3,i),
                        "engine_api",
                        "created_by_nodejs"
                      ]
                    }
              		}
              	};
                doc.createDimension(param);
              }
            });
          });
          break;
        case 'Load Variables':
          this.session.open().then((global) => {
            global.openDoc(AppID).then((doc) => {
              console.log('Loading ' + excel.data.meta.nrows + ' lines.');
              for(i = 0; i < excel.data.meta.nrows; i++) {
                var param = {
              		"qProp": {
              			"qInfo": {
              				"qId": "",
              				"qType": "variable"
              			},
              			"qMetaDef": {},
              			"qName": "" + excel.data.element(1,i),
              			"qComment": "" + excel.data.element(3,i),
              			"qNumberPresentation": {
              				"qType": 0,
              				"qnDec": 0,
              				"qUseThou": 0,
              				"qFmt": "",
              				"qDec": "",
              				"qThou": ""
              			},
              			"qIncludeInBookmark": false,
              			"qDefinition": "" + excel.data.element(2,i),
                    "tags": [
                      "ID" + excel.data.element(0,i),
                      "engine_api",
                      "created_by_nodejs"
                    ]
              		}
              	};
                doc.createVariableEx(param);
              }
            });
          });
          break;
		case 'Backup':
          this.session.open().then((global) => {
            global.openDoc(AppID).then((doc) => {
              var param = {
          			"qInfo": {
          				"qType": "MeasureList"
          			},
          			"qMeasureListDef": {
          				"qType": "measure",
          				"qData": {
          					"title": "/title",
          					"tags": "/tags"
          				}
          			}
          		};
              doc.createSessionObject(param)
              .then((mList) => {
                mList.getLayout()
                .then((layout) => {
                  QIX.listm = [];
                  for(i in layout.qMeasureList.qItems) {
                    QIX.listm.push(layout.qMeasureList.qItems[i].qInfo.qId);
                  }
                })
                .then(() => {
                  QIX.recordsM = [];
                  var it = 1;
                  for(i in QIX.listm) {
                    doc.getMeasure({"qId": QIX.listm[i]})
                    .then((measure) => {
                      measure.getLayout()
                      .then((layout) => {
                        QIX.recordsM.push({
                            c1: layout.qInfo.qId,
                            c2: layout.qMeta.title,
                            c3: layout.qMeasure.qLabelExpression,
                            c4: layout.qMeasure.qDef,
                            c5: layout.qMeta.description,
                            c6: layout.qMeta.tags.join()
                          });
                      })
                    })
                  }
                })
              });
            });
          });
          this.session.open().then((global) => {
            global.openDoc(AppID).then((doc) => {
              var param = {
          			"qInfo": {
          				"qType": "DimensionList"
          			},
          			"qDimensionListDef": {
          				"qType": "dimension",
          				"qData": {
          					"title": "/title",
          					"tags": "/tags",
          					"grouping": "/qDim/qGrouping",
          					"info": "/qDimInfos"
          				}
          			}
          		};
              doc.createSessionObject(param)
              .then((dList) => {
                dList.getLayout()
                .then((layout) => {
                  QIX.listd = [];
                  for(i in layout.qDimensionList.qItems) {
                    QIX.listd.push(layout.qDimensionList.qItems[i].qInfo.qId);
                  }
                })
                .then(() => {
                  QIX.recordsD = [];
                  var it = 1;
                  for(let i = 0 ; i < QIX.listd.length ; i++) {
                    doc.getDimension({"qId": QIX.listd[i]})
                    .then((dimension) => {
                      dimension.getLayout()
                      .then((layout) => {
                        QIX.recordsD.push({
                            c1: layout.qInfo.qId,
                            c2: layout.qMeta.title,
                            c3: layout.qDim.qFieldDefs,
                            c4: layout.qMeta.description,
                            c5: layout.qDim.qLabelExpression,
                            c6: layout.qMeta.tags.join()
                          });
                      })
                    })
                  }
                })
              });
            });
          });
          this.session.open().then((global) => {
            global.openDoc(AppID).then((doc) => {
              var param = {
        				"qInfo": {
        					"qType": "VariableList"
        				},
        				"qVariableListDef": {
        					"qType": "variable",
        					"qShowReserved": true,
        					"qShowConfig": true,
        					"qData": {
        						"tags": "/tags"
        					}
        				}
        			};
              doc.createSessionObject(param)
              .then((vList) => {
                vList.getLayout()
                .then((layout) => {
                  QIX.listv = [];
                  QIX.recordsV = [];
                  var it = 1;
                  for(i in layout.qVariableList.qItems) {
                    if(!("qIsReserved" in layout.qVariableList.qItems[i])) {
                      QIX.recordsV.push({
                          c1: layout.qVariableList.qItems[i].qInfo.qId,
                          c2: layout.qVariableList.qItems[i].qName,
                          c3: layout.qVariableList.qItems[i].qDefinition,
                          c4: layout.qVariableList.qItems[i].qDescription
                        });
                      QIX.listv.push(layout.qVariableList.qItems[i].qInfo.qId);
                    }
                  }
                });
              });
            });
          });
          break;
        case 'Create Adhoc Table':
          this.session.open().then((global) => {
            global.openDoc(AppID).then((doc) => {
              if(this.show == 1) {
                console.log('Creating Adhoc Table.');
              }
              const table = {
                "qProp" : {
                  "qInfo": {
                    "qId": "",
                    "qType": "masterobject"
                  },
                  "qMetaDef": {
                    "title": "Adhoc Table",
                    "description": "A table containing all master measures and dimensions, with conditional columns.",
                    "tags": [
                      "adhoc",
                      "engine_api",
                      "created_by_nodejs"
                    ]
                  },
                  "qHyperCubeDef": {
                    "qDimensions": [],
                    "qMeasures": [],
                    "qInterColumnSortOrder": [], //0-n
                    "qSuppressMissing": true,
                    "qInitialDataFetch": [],
                    "qReductionMode": "N",
                    "qMode": "S",
                    "qPseudoDimPos": -1,
                    "qNoOfLeftDims": -1,
                    "qMaxStackedCells": 5000,
                    "qCalcCond": {
                      "qv": "if(GetSelectedCount([_Dim]) + GetSelectedCount([_Ind]) > 0, 1, 0)"
                    },
                    "qTitle": {},
                    "qCalcCondition": {
                      "qCond": {
                        "qv": "if(GetSelectedCount([_Dim]) + GetSelectedCount([_Ind]) > 0, 1, 0)"
                      },
                      "qMsg": {
                        "qv": "Selecione pelo menos uma Dimensão ou Indicador para visualizar os dados."
                      }
                    },
                    "qColumnOrder": [], //0-n
                    "columnOrder": [], //0-n
                    "columnWidths": [], //-1
                    "qLayoutExclude": {
                      "qHyperCubeDef": {
                        "qDimensions": [],
                        "qMeasures": [],
                        "qInterColumnSortOrder": [],
                        "qInitialDataFetch": [],
                        "qReductionMode": "N",
                        "qMode": "S",
                        "qPseudoDimPos": -1,
                        "qNoOfLeftDims": -1,
                        "qMaxStackedCells": 5000,
                        "qCalcCond": {},
                        "qTitle": {},
                        "qCalcCondition": {
                          "qCond": {},
                          "qMsg": {}
                        },
                        "qColumnOrder": []
                      }
                    },
                    "customErrorMessage": {
                      "calcCond": "Selecione pelo menos uma Dimensão ou Indicador para visualizar os dados."
                    }
                  },
                  "showTitles": true,
                  "title": "Adhoc table",
                  "subtitle": "",
                  "footnote": "",
                  "showDetails": false,
                  "totals": {
                    "show": true,
                    "position": "noTotals",
                    "label": "Totais"
                  },
                  "scrolling": {
                    "keepFirstColumnInView": false
                  },
                  "multiline": {
                    "wrapTextInHeaders": true,
                    "wrapTextInCells": true
                  },
                  "visualization": "table",
                  "labelExpression": {
                    "qStringExpression": {
                      "qExpr": "'Adhoc table 0'"
                    }
                  },
                  "masterVersion": 0.96
                }
              };
              var arr1 = [], arr2 = [];
              for(i in QIX.DimensionList) {
                table.qProp.qHyperCubeDef.qDimensions.push(addDim(QIX.DimensionList[i].ID, QIX.DimensionList[i].Título));
                arr1.push(arr1.length);
                arr2.push(-1);
              }
              for(i in QIX.MeasureList) {
                table.qProp.qHyperCubeDef.qMeasures.push(addInd(QIX.MeasureList[i].ID, QIX.MeasureList[i].Título));
                arr1.push(arr1.length);
                arr2.push(-1);
              }
              table.qProp.qHyperCubeDef.qInterColumnSortOrder = arr1;
              table.qProp.qHyperCubeDef.qColumnOrder = arr1;
              table.qProp.qHyperCubeDef.columnOrder = arr1;
              table.qProp.qHyperCubeDef.columnWidths = arr2;
              setTimeout(() => doc.createObject(table).then(() => console.log("Done")).catch(() => console.log("Error")), 2000);
            });
          });
          break;
        default:
          console.log('\nUnsupported command\n');
          break;
      }
    }
  },
  listcontrols: function() {
    console.log('\n Exixting commands ');
    console.log(' 1  - List Measures ');
    console.log(' 2  - List Dimensions ');
    console.log(' 3  - List Variables ');
    console.log(' 4  - Delete all Measures ');
    console.log(' 5  - Delete all Dimensions ');
    console.log(' 6  - Delete all Variables ');
    console.log(' 7  - Load Measures (requires excel sheet previously loaded) ');
    console.log(' 8  - Load Dimensions (requires excel sheet previously loaded) ');
    console.log(' 9  - Load Variables (requires excel sheet previously loaded) ');
	console.log(' 10 - Backup all master Measures, master Dimensions and Variables ');
	console.log(' 11 - Create Adhoc Table with all master Measures and master Dimensions \n');
  }
};

function initialize() {
  process.stdout.write('\033c');
  console.log('\n\n       Welcome!\n');
  console.log(`   Anytime press 'm' to open Menu and 0 to exit.\n`);
  rl.prompt();
};

rl.on('line', (line) => {
  if(line !== null) {
    menuHandler(line.trim());
  }
})
.on('close', () => {
  console.log('\n Have a great day!');
  process.exit(0);
});

const menu = () => {
  process.stdout.write('\033c');
  console.log('         Menu          \n\n');
  console.log('   1 - Open Excel file' + (excel.init > 0 ? ('    -  File ' + excel.file + ': Ok') : '') + (excel.init == 2 ? ('  ;  Sheet ' + excel.currSheet + ' open') : ''));
  console.log('   2 - Connect to QS - QIX Engine ' + ("AppID" in QIX ? ('  -  Connected to app: ' + QIX.AppName) : ''));
  console.log('   0 - Exit\n');
  rl.prompt();
  menuHandler = function(input){
    switch(input) {
      case 'm': menu(); break;
      case '1': menuExcel(); break;
      case '2': menuQS(); break;
      case '0':
        console.log('\n Have a great day!!');
        process.exit(0);
        break;
      case '':
        rl.prompt();
        break;
      default:
        console.log("'" + input + "' is not a known command.");
        rl.prompt();
    }
  };
};

const menuExcel = () => {
  process.stdout.write('\033c');
  console.log('         Menu - Excel         \n\n');
  if(excel.init == 2) {
    console.log("Sheet " + excel.currSheet + ", of file " + excel.file + ", open. Do you want to keep it open or change the sheet?\n")
    console.log('1 - Keep');
    console.log('2 - Change sheet');
    console.log('0 - Change file\n');
    console.log('prev(n) - Preview of current sheet; n is the number of lines to be previewed (default: n = 10)\n');
  } else if (excel.init == 1) {
    console.log("File " + excel.file + " open. Do you want to pick a sheet or change the file?");
    console.log('1 - Choose sheet');
    console.log('2 - Change file');
    console.log('0 - Exit\n');
  } else {
    console.log('Enter the name of the excel file: ');
  }

  rl.prompt();
  menuHandler = function(input) {
    switch(input) {
      case '1':
        if(excel.init == 2) {
          menu();
        } else if(excel.init == 1) {
          menuExcelsheet();
        } else {
          rl.prompt();
        }
        break;
      case '2':
        if(excel.init == 2) {
          menuExcelsheet();
        } else if(excel.init == 1) {
          excel.reset();
          menuExcel();
        } else {
          rl.prompt();
        }
        break;
      case '0':
        if(excel.init == 2) {
          excel.reset();
          menuExcel();
        } else {
          console.log('\n Have a nice day!!');
          process.exit(0);
        }
        break;
      case 'm':
        menu();
        break;
      case (input.match(/^prev/) || {}).input:
        excel.preview(input.substring(5, input.length - 1));
        rl.prompt();
        break;
      case '':
        rl.prompt();
        break;
      default:
        if(excel.init > 0) {
          menuExcel();
        }
        if(input.substring(input.length - 5, input.length) == '.xlsx') {
          if(fs.existsSync(input)) {
            excel.open(input);
          }
        } else if(fs.existsSync((input == '' ? ' ' : input) + ".xlsx")) {
          excel.open(input + ".xlsx");
        }
        if(excel.init == 0) {
          console.log("Invalid file. Try again");
          rl.prompt();
        } else {
          menuExcelsheet();
        }
    }
  };
};

const menuExcelsheet = () => {
  process.stdout.write('\033c');
  console.log('         Menu - Excel         \n\n');
  excel.listSheets();
  rl.prompt();
  menuHandler = function(input){
        if(Number(input) > 0 && Number(input) <= excel.nsheets) {
          excel.setSheet(Number(input) - 1);
          menuExcel();
        } else if(Number(input) == 0) {
          menuExcelsheet();
        } else if(input == 'm') {
          menuExcel();
        } else if(input == '') {
          rl.prompt();
        } else {
          console.log('Invalid input. Enter 0 to go back.');
        }
        rl.prompt();
    };
};

const menuQS = () => {
  process.stdout.write('\033c');
  console.log('         Menu - QIX          \n\n');
  if("AppID" in QIX) {
    console.log('Connected to app: ' + QIX.AppName + '. Do you want to: ');
    console.log('   1 - Keep connected to this app');
    console.log('   2 - Change app');
    console.log('   0 - Exit\n');
  } else {
    QIX.ini();
    console.log('\n    Enter the AppID (Server) or App file path (Desktop)\n');
  }
  rl.prompt();
  menuHandler = function(input) {
    switch (input) {
      case '1':
        menuQScom();
        break;
      case '2':
        delete QIX.AppID;
        menuQS();
        break;
      case '0':
        console.log('\n Have a nice day!!');
        process.exit(0);
        break;
      case 'm':
        menu();
        break;
      case '':
        rl.prompt();
        break;
      default:
        if(input.length > 0/*== 36*/){
          QIX.try(input);
          setTimeout(function() {("AppID" in QIX ? menuQScom() : menuQS())}, 2500);
        } else {
          console.log('Invalid AppID');
          rl.prompt();
        }
    }
  };
};

const menuQScom = () => {
  process.stdout.write('\033c');
  console.log('         Menu - QIX          \n');
  QIX.listcontrols();
  rl.prompt();
  menuHandler = function(input) {
    switch (input) {
      case '1':
        QIX.on('List Measures');
        setTimeout(function() {rl.prompt()}, 500);
        break;
      case '2':
        QIX.on('List Dimensions');
        setTimeout(function() {rl.prompt()}, 500);
        break;
      case '3':
        QIX.on('List Variables');
        setTimeout(function() {rl.prompt()}, 500);
        break;
      case '4':
        QIX.on('Delete All Measures');
        setTimeout(function() {rl.prompt()}, 500);
        break;
      case '5':
        QIX.on('Delete All Dimensions');
        setTimeout(function() {rl.prompt()}, 500);
        break;
      case '6':
        QIX.on('Delete All Variables');
        setTimeout(function() {rl.prompt()}, 500);
        break;
      case '7':
        if(excel.init == 2) {
          QIX.on('Load Measures');
        } else {
          console.log('No sheet loaded!!! Go back to the Menu and load the appropriate sheet.');
        }
        setTimeout(function() {rl.prompt()}, 500);
        break;
      case '8':
        if(excel.init == 2) {
          QIX.on('Load Dimensions');
        } else {
          console.log('No sheet loaded!!! Go back to the Menu and load the appropriate sheet.');
        }
        setTimeout(function() {rl.prompt()}, 500);
        break;
      case '9':
        if(excel.init == 2) {
          QIX.on('Load Variables');
        } else {
          console.log('No sheet loaded!!! Go back to the Menu and load the appropriate sheet.');
        }
        setTimeout(function() {rl.prompt()}, 500);
        break;
	  case '10':
        QIX.on('Backup');
        setTimeout(function() {rl.prompt()}, 500);
        break;
      case '11':
        QIX.on('Create Adhoc Table');
        setTimeout(function() {rl.prompt()}, 500);
        break;
      case '0':
        console.log('\n Have a nice day!!');
        process.exit(0);
        break;
      case 'm':
        menu();
        break;
      case '':
        rl.prompt();
        break;
      default:
        console.log('Unknown command');
        rl.prompt();
    }
  };
};

initialize();
