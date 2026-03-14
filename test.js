// ==UserScript==
// @name         Youtube Right Click Popup
// @namespace    http://tampermonkey.net/
// @version      0.1
// @description  try to take over the world!
// @author       You
// @match        *://www.youtube.com/*
// @match        *://youtube.com/*
// @match        *://youtube.com
// @match        *://www.youtube.com
// @icon     https://i.pinimg.com/originals/31/23/9a/31239a2f70e4f8e4e3263fafb00ace1c.png
// @run-at document-start
// @grant  GM_addStyle
// @grant  GM_openInTab
// ==/UserScript==

var url = window.location.href;
const channel_name = ['中天新聞','哏傳媒', '全球大視野', '香港01','飛碟聯播網',
                      '頭條開講','大新聞大爆卦','少康戰情室 TVBS Situation Room',
                      'Muse木棉花-HK','新聞大白話','即新聞','中視新聞',
                     ]
//GM_addStyle (".yt-core-attributed-string--link-inherit-color {color: #000;}");
//GM_addStyle ("ytd-shorts #shorts-container.ytd-shorts { display: unset; justify-content: center; height: calc(100vh - var(--ytd-shorts-masthead-height));overflow-x: hidden; overflow-y: scroll;    scroll-snap-type: y mandatory; scroll-padding-top: 0; margin-top: var(--ytd-shorts-top-margin-free-scroll-override, var(--ytd-shorts-top-spacing)); scrollbar-width: none; -ms-overflow-style: none;}");
//GM_addStyle ("#shorts-player #immersive-translate-caption-window {height: 85%;}");
//GM_addStyle (".ytd-thumbnail-overlay-time-status-renderer {display: inline-block;position: absolute;top: 0px;left: 0px;margin: 4px;display: flexbox;display: flex;flex-direction: row;}");

window.addEventListener('contextmenu', function (e) {
  //console.log(e.target);
  let ppp = e.target.parentNode.parentNode.parentNode
  //console.log(ppp.outerHTML)
  //console.log(ppp.className)
  var timestamp = null;
  var starttime = null;
  if (ppp.outerHTML.match('首播日期：') != null){
      var starttime = ppp.outerHTML.match('首播日期：(.*?)<')[1].split(' ')
      const [day, month, year] = starttime[0].split('/')
      const [hour, minute] = starttime[1].split(':')
      //console.log(year, month, day)
      const dateObj = new Date(year, month-1, day, hour, minute)
      timestamp = dateObj.getTime()/1000
  }
  if (ppp.outerHTML.match('預定發佈時間：') != null){
      var starttime = ppp.outerHTML.match('預定發佈時間：(.*?)<')[1].split(' ')
      const [day, month, year] = starttime[0].split('/')
      const [hour, minute] = starttime[1].split(':')
      //console.log(year, month, day)
      const dateObj = new Date(year, month-1, day, hour, minute)
      timestamp = dateObj.getTime()/1000
  }
  //console.log(e.target.parentNode.parentNode.href)
  //console.log(e.target.style);
  //alert(e.target);
  let info_title = e.target.getElementsByClassName("ytp-videowall-still-info-title");
  let pic_tran = e.target.style;
    //style-scope ytd-thumbnail-overlay-toggle-button-renderer
  //console.log();
  //console.log(e.target.outerHTML.match('polyline'));

  //console.log(pic_tran.cssText)
  
  if (pic_tran.cssText == 'background-color: transparent;'){
      console.log(e.target.parentNode.parentNode.href.replace('https', 'yurl'))
      window.open(e.target.parentNode.parentNode.href.replace('https', 'yurl'))
      /*
      //console.log(e.target.parentNode.parentNode.href)
      console.log ('BSAFD')
      var link = "https://www.youtube.com/embed/"+e.target.src.match('\/\/.*\/(.*?)\/')[1]+"?autoplay=1";
      window.open(link,'YoutubepopUpWindow'+Date.now(),'height=253,width=450,left=1470,top=50,resizable=no,scrollbars=yes,toolbar=yes,menubar=no,location=no,directories=no ');
      */
  }
  else if (e.target.id == 'img'){
      console.log('AAAA');
      //console.log(e.target.parentNode.parentNode.parentNode.parentNode.outerHTML.match('badge badge-style-type-live-now style-scope ytd-badge-supported-renderer'));
      //e.preventDefault();
      console.log(e.target.src);
      //console.log(e.target.src.match('\/\/.*\/(.*?)\/')[1]);
      /*
      var link = "https://www.youtube.com/embed/"+e.target.src.match('\/\/.*\/(.*?)\/')[1]+"?autoplay=1";
      var linkb = "https://www.youtube.com/live_chat?is_popout=1&v="+e.target.src.match('\/\/.*\/(.*?)\/')[1];
      //console.log(link);
      //window.open("'"+link+"'",'height=253,width=450,left=1470,top=50,resizable=no,scrollbars=yes,toolbar=yes,menubar=no,location=no,directories=no ');
      window.open(link,'YoutubepopUpWindow'+Date.now(),'height=253,width=450,left=1470,top=50,resizable=no,scrollbars=yes,toolbar=yes,menubar=no,location=no,directories=no ');
      if (e.target.parentNode.parentNode.parentNode.parentNode.outerHTML.match('badge badge-style-type-live-now style-scope ytd-badge-supported-renderer') != null){
          window.open(linkb,'YoutubepopUpChat'+Date.now(),'height=450,width=320,left=1470,top=350,resizable=no,scrollbars=yes,toolbar=yes,menubar=no,location=no,directories=no ');
      }
      */
  }
  else if ((info_title != null) && (url.match('autoplay') != null)){
      console.log('ABC');
      let linkc = e.target.parentNode.parentNode.parentNode.parentNode.href;
      let linkd = e.target.parentNode.parentNode.parentNode.href
      if (linkc != null){
          console.log('ABC1');
          //console.log(linkc);
          e.preventDefault();
          window.open(linkc.replace('https://www.youtube.com/watch?v=', 'https://www.youtube.com/embed/').replace('&feature=emb_rel_end', '?autoplay=1'),"_self")
      }
      else if (linkd != null){
          console.log('ABC2');
          //console.log(e.target.parentNode);
          //console.log(e.target.parentNode.parentNode);
          //console.log(e.target.parentNode.parentNode.parentNode);
          //console.log(linkd);
          e.preventDefault();
          window.open(linkd.replace('https://www.youtube.com/watch?v=', 'https://www.youtube.com/embed/').replace('&feature=emb_rel_end', '?autoplay=1'),"_self")
      }
  }
  else if (e.target.id == 'thumbnail'){
      console.log('CCC');
      e.preventDefault();
      console.log(e.target.parentNode.parentNode);
      //console.log(e.target.src.match('\/\/.*\/(.*?)\/')[1]);
      var linkb = "https://www.youtube.com/embed/"+e.target.src.match('\/\/.*\/(.*?)\/')[1]+"?autoplay=1";
      //console.log(linkb);
      window.open(linkb,'YoutubepopUpWindow','height=253,width=450,left=1470,top=50,resizable=no,scrollbars=yes,toolbar=yes,menubar=no,location=no,directories=no ');
  }
  else if (e.target.id == 'video-title'){
      //console.log('CCCdddd');
      e.preventDefault();
      //console.log(e.target.parentNode.title);
      if (timestamp != null){
          new Notification("在"+starttime+'將會播放', {"body": e.target.parentNode.title,});
          //new Notification("This is a test notification", {"body": "You have allowed notifications, this is just a test",});
          window.open(e.target.parentNode.href.replace('https', 'yurl').replace('shorts/', 'watch?v=')+'&start_time='+timestamp)
      }
      else{
          window.open(e.target.parentNode.href.replace('https', 'yurl').replace('shorts/', 'watch?v='))
      }
      /*
      //console.log(e.target.src.match('\/\/.*\/(.*?)\/')[1]);
      var linkb = "https://www.youtube.com/embed/"+e.target.src.match('\/\/.*\/(.*?)\/')[1]+"?autoplay=1";
      //console.log(linkb);
      window.open(linkb,'YoutubepopUpWindow','height=253,width=450,left=1470,top=50,resizable=no,scrollbars=yes,toolbar=yes,menubar=no,location=no,directories=no ');
      */
  }

  else if (e.target.parentNode.parentNode.id == 'inline-preview-player'){
      console.log('GDSFSDGD')
      //console.log(e.target.parentNode.parentNode.getElementsByClassName('ytp-title-link yt-uix-sessionlink').href);
      //e.preventDefault();
      //console.log(e.target.src);
      //console.log(e.target.src.match('\/\/.*\/(.*?)\/')[1]);
      var nlink = e.target.parentNode.parentNode.getElementsByClassName('ytp-title-link yt-uix-sessionlink')[0].href
      //var link = "https://www.youtube.com/embed/"+e.target.src.match('\/\/.*\/(.*?)\/')[1]+"?autoplay=1";
      //var link = "https://www.youtube.com/embed/"+nlink.match('watch\\?v=(.*)')[1]+"?autoplay=1";
      //window.open("'"+link+"'",'height=253,width=450,left=1470,top=50,resizable=no,scrollbars=yes,toolbar=yes,menubar=no,location=no,directories=no ');
      //window.open(link,'YoutubepopUpWindow','height=253,width=450,left=1470,top=50,resizable=no,scrollbars=yes,toolbar=yes,menubar=no,location=no,directories=no ');
      window.open(nlink.replace('https', 'yurl'));
  }
  else if (e.target.parentNode.parentNode.id == 'thumbnail'){
      //console.log(e.target.parentNode.parentNode.getElementsByClassName('ytp-title-link yt-uix-sessionlink').href);
      //e.preventDefault();
      console.log('DDDDDD');
      window.open(e.target.parentNode.parentNode.href.replace('https', 'yurl'))
      /*
      //console.log(e.target.src);
      //console.log(e.target.src.match('\/\/.*\/(.*?)\/')[1]);
      //var nlink = e.target.parentNode.parentNode.href;//.split('&')[0];
      var nlink = "https://www.youtube.com/embed/"+e.target.parentNode.parentNode.href.match('watch\\?v=(.*)')[1].split('&')[0]+"?autoplay=1";
      window.open(nlink,'YoutubepopUpWindow','height=253,width=450,left=1470,top=50,resizable=no,scrollbars=yes,toolbar=yes,menubar=no,location=no,directories=no ');
      //console.log(nlink);
      */
      /*
      var nlink = e.target.parentNode.parentNode.getElementsByClassName('ytp-title-link yt-uix-sessionlink')[0].href
      //var link = "https://www.youtube.com/embed/"+e.target.src.match('\/\/.*\/(.*?)\/')[1]+"?autoplay=1";
      var link = "https://www.youtube.com/embed/"+nlink.match('watch\\?v=(.*)')[1]+"?autoplay=1";
      //window.open("'"+link+"'",'height=253,width=450,left=1470,top=50,resizable=no,scrollbars=yes,toolbar=yes,menubar=no,location=no,directories=no ');
      window.open(link,'YoutubepopUpWindow','height=253,width=450,left=1470,top=50,resizable=no,scrollbars=yes,toolbar=yes,menubar=no,location=no,directories=no ');
      */
  }
  else if (ppp.className == "yt-lockup-view-model__content-image"){
      //console.log(e.href)
      window.open(ppp.href.replace('https', 'yurl'))
  }
  //else if (e.target.outerHTML.match('polyline') !=null){
  //    console.log(e.target.points);
  //}
}, false);

/*
if (url.match('rainmeter') != null){
    alert('ABC');
    url = url.replace('https', 'yurl').replace('&rainmeter', '');
    window.open(e.target.parentNode.parentNode.href.replace('https', 'yurl'))
    
    ///url = url.replace('https://www.youtube.com/watch\?v=', 'https://www.youtube.com/embed/');
    ///window.open(url.replace('&rainmeter', '?autoplay=1'),'YoutubepopUpWindow','height=253,width=450,left=1470,top=50,resizable=no,scrollbars=yes,toolbar=yes,menubar=no,location=no,directories=no ');
    ///window.close();
    //setTimeout(function(){window.close(); }, 500);
    
}
