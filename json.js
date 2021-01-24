		function parseJson(json_str) {
			// JSON ��������I�u�W�F�N�g��
			var json = eval("(" + json_str + ")");
			return json;
		}
		// ���I�ȃp�X�Ƃ��Ēl�ɃA�N�Z�X���܂�
		function getJson(json, arr_vb_obj) {

			// arr_vb_obj �́AVBArray(�Z�[�t�z��)
			// arr_vb �́AJScript ���ł� VBArray���b�p�[
			var arr_vb = new VBArray(arr_vb_obj);

			// JScript �̔z��ɕϊ�
			var arr = arr_vb.toArray();

			// �K�w�\���� JSON �������B�肾������
			// ���̏ꍇ�͌��ʂƂ���  "�f�[�^" �����o��
			var trav = json;
			for( var i = 0; i < arr.length; i++ ) {
				trav = trav[arr[i]]
			}

			return trav;
		}
		// �c���[�Ƃ��ă_���v���܂�
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
		// XML �Ƃ��ă_���v���܂�
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
