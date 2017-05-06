var exec = require("child_process").exec;

var manifest = require("../vss-extension.json");
var extensionId = manifest.id;

// Package extension
var command = `tfx extension create --overrides-file configs/dev.json --manifest-globs vss-extension.json --rev-version --extension-id ${extensionId}-dev --no-prompt --json`;
exec(command, (error, stdout) => {
    if (error) {
        console.error(`Could not create package: '${error}'`);
        return;
    }

    console.log(`Package created`);
});