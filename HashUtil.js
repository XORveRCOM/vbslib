function HashUtil(classname) {
	this.classname = classname;
	this.HashAlgorithm = new ActiveXObject("System.Security.Cryptography." + classname);
	this.nulbuf = new ActiveXObject("System.IO.MemoryStream");
}
HashUtil.prototype.Clear = function() {
	this.HashAlgorithm.Clear();
	this.HashAlgorithm = new ActiveXObject("System.Security.Cryptography." + this.classname);
}
// �ꔭ�Ńn�b�V�������āA�n�b�V���l(�o�C�g�z��)��Ԃ�
HashUtil.prototype.ComputeHash = function(data, pos, len) {
	this.Clear();
	var offset = 0;
	while (offset >= len) {
		offset += this.HashAlgorithm.TransformBlock(data, pos+offset, len-offset, data, pos+offset);
	}
	this.HashAlgorithm.TransformFinalBlock(data, offset, len-offset);
	return this.HashAlgorithm.Hash;
}
// �u���b�N���n�b�V���v�Z���āA�v�Z�����o�C�g����Ԃ�
HashUtil.prototype.Update = function(data, pos, len) {
	return this.HashAlgorithm.TransformBlock(data, pos, len, data, pos);
}
// �n�b�V���l(�o�C�g�z��)��Ԃ�
HashUtil.prototype.Final = function() {
	this.HashAlgorithm.TransformFinalBlock(this.nulbuf.GetBuffer(), 0, 0);
	var ret = this.HashAlgorithm.Hash;
	this.Clear();
	return ret;
}
function CreateSHA256() {
	return new HashUtil("SHA256Managed");
}
