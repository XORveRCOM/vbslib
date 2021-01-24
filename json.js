		function parseJson(json_str) {
			// JSON 文字列をオブジェクト化
			var json = eval("(" + json_str + ")");
			return json;
		}
		// 動的なパスとして値にアクセスします
		function getJson(json, arr_vb_obj) {

			// arr_vb_obj は、VBArray(セーフ配列)
			// arr_vb は、JScript 内での VBArrayラッパー
			var arr_vb = new VBArray(arr_vb_obj);

			// JScript の配列に変換
			var arr = arr_vb.toArray();

			// 階層構造の JSON を順次撮りだす処理
			// この場合は結果として  "データ" を取り出す
			var trav = json;
			for( var i = 0; i < arr.length; i++ ) {
				trav = trav[arr[i]]
			}

			return trav;
		}
		// ツリーとしてダンプします
		function json2tree(json, indent) {
			var res = "";
			for (var name in json) {
				var val = json[name];
				var vartype = typeof(val);
				switch (vartype) {
				case "object":
					res = res + indent + name + "\n";
					res = res + json2tree(val, indent + "    ");
					break;
				case "string":
					res = res + indent + name + " (" + vartype + ") : \"" + val + "\"\n";
					break;
				case "number":
				case "boolean":
					res = res + indent + name + " (" + vartype + ") : " + val + "\n";
					break;
				default :
				}
			}
			return res;
		}
		// XML としてダンプします
		function json2xml(json, indent) {
			var res = "";
			for (var name in json) {
				var val = json[name];
				var vartype = typeof(val);
				switch (vartype) {
				case "object":
					res = res + indent + "<node name=\"" + name + "\">\n";
					res = res + json2xml(val, indent + "  ");
					res = res + indent + "</node>\n";
					break;
				case "string":
				case "number":
				case "boolean":
				default :
					res = res + indent + "<item name=\"" + name + "\" type=\"" + vartype + "\" value=\"" + val + "\"/>\n";
					break;
				}
			}
			return res;
		}
		function IsArray(array) {
			return !(
				!array || 
				(!array.length || array.length == 0) || 
				typeof array !== 'object' || 
				!array.constructor || 
				array.nodeType || 
				array.item 
			);
		}
