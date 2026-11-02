const {chromium}=require('playwright');
(async()=>{const b=await chromium.launch();const p=await b.newPage({viewport:{width:1920,height:1080},deviceScaleFactor:2});
await p.goto('file://'+__dirname+'/timeline.html');await p.screenshot({path:__dirname+'/journal-timeline.png'});await b.close();})();
