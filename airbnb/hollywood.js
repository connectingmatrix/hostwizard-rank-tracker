const spawn = require("child_process").spawn;
console.log('Running script')

console.log('Starting..')
const pythonProcess = spawn('python',["./scrapper-cron-hollywood.py", 'runcrons'])
console.log('Started')

pythonProcess.stdout.on('data', async (data) => {
    console.log(data.toString())
})

pythonProcess.stderr.on('data', async (data) => {
    console.error(data.toString())
})

pythonProcess.on('exit', ()=>{
    process.exit(1)
})
pythonProcess.on('close', ()=>{
    process.exit(1)
})