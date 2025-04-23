#!/usr/bin/env node
const program = require('commander');

//Actions

const uninstallPackage = require('./actions/uninstall-package.action')
const buildPackage = require('./actions/build-package.action')
const listPackages = require('./actions/list-packages.action')
const createPackageScaffold = require('./actions/create-package.action');
const installPackage = require('./actions/install.action')
const addVbaFile = require('./actions/add-file.action')


//Install
program.command('install').description('xvba install [package] - Instal Xvba Package from Xvba Repository').action((package) => { installPackage(package) })
program.command('i').description('xvba i [package] - Instal Xvba Package from Xvba Repository').action((package) => { installPackage(package) })
//Uninstall
program.command('uninstall').description('xvba uninstall [package] -  Uninstall Package').action((package) => { uninstallPackage(package) });

program.command('create').description('xvba create [package] - Create Xvba Package scaffold on xvba_modules').action((package) => { createPackageScaffold(package) })


program.command('build').description('xvba build [package] - Build Xvba Package file ').action((package) => { buildPackage(package) })

//Add new vba files
program.command('add [m,c]').description('xvba add -c [sub-folder/file] or xvba add -m [sub-folder/file] - Create Vba file type module bas or class cls').action( async (filePath) => { addVbaFile(filePath,program) })

program.command('export')
  .description('xvba export - Export VBA modules from Excel file')
  .action(() => {
    const fs = require('fs');
    const path = require('path');

    // config.json を読み込む
    const configPath = path.join(__dirname, 'config.json');
    if (!fs.existsSync(configPath)) {
      console.error('Error: config.json not found. Please run "xvba init" first.');
      process.exit(1);
    }

    const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
    const excelFile = config.excel_file;

    if (!excelFile || !fs.existsSync(excelFile)) {
      console.error(`Error: Excel file "${excelFile}" not found.`);
      process.exit(1);
    }

    console.log(`Exporting VBA modules from ${excelFile}...`);
    // エクスポート処理をここに記述
    // 例: VBA モジュールをエクスポートして xva_modules フォルダに保存
  });

program.command('init')
  .description('xvba init - Create a new config files')
  .action(() => {
    console.log('Initializing XVBA...');
    // 必要な初期化処理をここに記述
    // 例: config.json ファイルを作成する処理
    const fs = require('fs');
    const configContent = {
      app_name: "XVBA",
      description: "",
      author: "",
      email: "",
      create_date: new Date().toString(),
      excel_file: "example.xlsm",
      ribbon_file: "customUI14",
      ribbon_folder: "ribbons",
      logs: "on",
      xvba_packages: {},
      xvba_dev_packages: {}
    };

    fs.writeFileSync('config.json', JSON.stringify(configContent, null, 2));
    console.log('config.json has been created.');
  });

program.command('ls').description('xvba ls - List all Packages installed').action(() => { listPackages() });

program
  .option('-p, --package', 'Create a Package')
  .option('-m, --module', 'Create Module')
  .option('-c, --class', 'Create a Class');

program.parse(process.argv);
