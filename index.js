var childProcess = require('child_process');

module.exports = function (sheets) {
	return new Promise(function (resolve, reject) {
		var child = childProcess.fork(__dirname + '/generator');
		child.send({
			sheets: sheets
		});
		child.on('message', function (message) {
			if (message.err) {
				reject(message.err);
			} else {
				resolve(message.result);
			}
			child.kill();
		});
	});
};