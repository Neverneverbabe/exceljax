const { execSync } = require('child_process');

function run(command) {
  execSync(command, { stdio: 'inherit' });
}

const tasks = {
  install() {
    run('npm install');
  },
  build() {
    run('npm run build');
  },
  buildGhPages() {
    run('npm run build:ghpages');
  },
  proxy() {
    run('node proxy-server.js');
  },
  start() {
    run('npm start');
  },
};

const task = process.argv[2];
if (!task || !tasks[task]) {
  console.log('Usage: node scripts/tasks.js <task>');
  console.log('Available tasks: ' + Object.keys(tasks).join(', '));
  process.exit(1);
}

tasks[task]();
