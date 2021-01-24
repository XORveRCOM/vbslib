function HashUtil(classname) {
	this.classname = classname;
	this.HashAlgorithm = new ActiveXObject("System.Security.Cryptography." + classname);
	this.nulbuf = new ActiveXObject("System.IO.MemoryStream");
}
HashUtil.prototype.Clear = function() {
	this.HashAlgorithm.Clear();
	this.HashAlgorithm = new ActiveXObject("System.Security.Cryptography." + this.classname);
}
// 一発でハッシュ化して、ハッシュ値(バイト配列)を返す
HashUtil.prototype.ComputeHash = function(data, pos, len) {
	this.Clear();
	var offset = 0;
	while (offset >= len) {
		offset += this.HashAlgorithm.TransformBlock(data, pos+offset, len-offset, data, pos+offset);
	}
	this.HashAlgorithm.TransformFinalBlock(data, offset, len-offset);
	return this.HashAlgorithm.Hash;
}
// ブロックをハッシュ計算して、計算したバイト数を返す
HashUtil.prototype.Update = function(data, pos, len) {
	return this.HashAlgorithm.TransformBlock(data, pos, len, data, pos);
}
// ハッシュ値(バイト配列)を返す
HashUtil.prototype.Final = function() {
	this.HashAlgorithm.TransformFinalBlock(this.nulbuf.GetBuffer(), 0, 0);
	var ret = this.HashAlgorithm.Hash;
	this.Clear();
	return ret;
}
function CreateSHA256() {
	return new HashUtil("SHA256Managed");
}
