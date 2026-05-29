[index.html](https://github.com/user-attachments/files/28372610/index.html)
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<meta name="apple-mobile-web-app-title" content="PAPS 측정">
<meta name="theme-color" content="#ffffff">
<title>PAPS 측정</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@3.4.0/tabler-icons.min.css">
<style>
*{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%;overflow-x:hidden}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f5f5f0;color:#1a1a1a}
:root{
  --color-background-primary:#ffffff;
  --color-background-secondary:#f5f5f0;
  --color-background-tertiary:#ebebeb;
  --color-text-primary:#1a1a1a;
  --color-text-secondary:#6b6b6b;
  --color-text-tertiary:#9b9b9b;
  --color-border-secondary:#d4d4d0;
  --color-border-tertiary:#e8e8e4;
  --border-radius-md:8px;
  --border-radius-lg:12px;
}
.app{max-width:420px;margin:0 auto;padding-bottom:2rem;min-height:100vh}
.topbar{background:#fff;border-bottom:0.5px solid #e8e8e4;padding:10px 14px;display:flex;align-items:center;gap:8px;position:sticky;top:0;z-index:20}
.tb-back{width:30px;height:30px;border:none;background:transparent;cursor:pointer;color:#6b6b6b;font-size:18px;display:none;align-items:center;justify-content:center;flex-shrink:0}
.tb-title{font-size:15px;font-weight:500;color:#1a1a1a;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.tb-act{font-size:12px;padding:5px 9px;border:0.5px solid #d4d4d0;border-radius:8px;background:#f5f5f0;color:#6b6b6b;cursor:pointer;flex-shrink:0}
.admin-btn{display:none;align-items:center;gap:5px;font-size:11px;padding:4px 10px;border-radius:20px;background:#EAF3DE;color:#3B6D11;font-weight:500;border:none;cursor:pointer;flex-shrink:0}
.admin-btn:hover{background:#D4ECC6}
.screen{padding:12px;display:none}.screen.on{display:block}
.tab-bar{display:grid;grid-template-columns:1fr 1fr 1fr;border:0.5px solid #e8e8e4;border-radius:12px;overflow:hidden;margin-bottom:12px;background:#f5f5f0}
.tab{padding:9px 4px;text-align:center;font-size:12px;font-weight:500;color:#6b6b6b;cursor:pointer;border:none;background:transparent}
.tab.on{background:#fff;color:#1a1a1a}
.card{background:#fff;border:0.5px solid #e8e8e4;border-radius:12px;padding:14px;margin-bottom:10px}
.ct{font-size:12px;font-weight:500;color:#6b6b6b;margin-bottom:10px;display:flex;align-items:center;gap:5px}
.fg{display:flex;flex-direction:column;gap:4px;margin-bottom:8px}
.fg:last-child{margin-bottom:0}
label{font-size:12px;color:#6b6b6b}
input,select,textarea{width:100%;border:0.5px solid #d4d4d0;border-radius:8px;padding:0 10px;font-size:14px;background:#fff;color:#1a1a1a;font-family:inherit}
input,select{height:38px}
textarea{padding:8px 10px;resize:none}
input:focus,select:focus,textarea:focus{outline:none;border-color:#378ADD;box-shadow:0 0 0 2px rgba(55,138,221,.12)}
input[readonly]{background:#f5f5f0;color:#6b6b6b}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px}
.g4{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:6px}
.bp{width:100%;height:44px;background:#378ADD;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:500;cursor:pointer}
.bp:active{transform:scale(.98)}
.bs{width:100%;height:38px;background:transparent;color:#6b6b6b;border:0.5px solid #d4d4d0;border-radius:8px;font-size:13px;cursor:pointer}
.bs:hover{background:#f5f5f0}
.bsm{height:32px;padding:0 12px;border:0.5px solid #d4d4d0;border-radius:8px;background:transparent;color:#6b6b6b;font-size:12px;cursor:pointer}
.refresh-btn{width:32px;height:32px;border:0.5px solid #d4d4d0;border-radius:8px;background:transparent;cursor:pointer;display:flex;align-items:center;justify-content:center;color:#6b6b6b;font-size:15px}
.refresh-btn:hover{background:#f5f5f0}
@keyframes spin{from{transform:rotate(0)}to{transform:rotate(360deg)}}
.spinning{animation:spin .5s ease}
.dropzone{border:2px dashed #d4d4d0;border-radius:12px;padding:24px 16px;text-align:center;cursor:pointer;margin-bottom:10px}
.dropzone:hover,.dropzone.drag{border-color:#378ADD;background:#E6F1FB}
#xlsx_input{display:none}
.ur{border-radius:8px;padding:10px 12px;margin-bottom:10px;font-size:13px;display:none}
.ur.ok{background:#EAF3DE;color:#3B6D11;border:0.5px solid #9FE1CB}
.ur.err{background:#FCEBEB;color:#A32D2D;border:0.5px solid #F0997B}
.cr{display:flex;align-items:center;gap:8px;padding:10px 12px;background:#f5f5f0;border-radius:8px;margin-bottom:6px;cursor:pointer;border:0.5px solid #e8e8e4}
.cr:hover{background:#ebebeb}
.ep{display:inline-flex;align-items:center;gap:4px;font-size:11px;padding:3px 8px;border-radius:20px;font-weight:500}
.sn{background:#fff;border:0.5px solid #e8e8e4;border-radius:12px;padding:12px 14px;margin-bottom:10px}
.sti{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px}
.sid{display:flex;align-items:center;gap:10px}
.sci{width:44px;height:44px;border-radius:50%;background:#E6F1FB;color:#185FA5;display:flex;align-items:center;justify-content:center;font-size:18px;font-weight:500;flex-shrink:0}
.snm{font-size:17px;font-weight:500}
.ssb{font-size:12px;color:#6b6b6b;margin-top:1px}
.pb-box{background:#f5f5f0;border-radius:8px;padding:10px 14px;margin-bottom:10px;display:flex;align-items:center;justify-content:space-between}
.pbar{height:8px;background:#ebebeb;border-radius:4px;overflow:hidden}
.pbf{height:100%;background:#378ADD;border-radius:4px;transition:width .3s}
.pbls{display:flex;justify-content:space-between;margin-top:4px;font-size:11px;color:#9b9b9b}
.nr{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.nb{height:36px;border:0.5px solid #d4d4d0;border-radius:8px;background:transparent;color:#6b6b6b;font-size:13px;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:4px}
.nb:hover{background:#f5f5f0}
.etb{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:12px}
.et{padding:6px 12px;border-radius:20px;font-size:12px;font-weight:500;cursor:pointer;border:1.5px solid transparent;background:#f5f5f0;color:#6b6b6b}
.et.on{color:#fff}
.evb{padding:4px 0}
.evh{display:flex;align-items:center;gap:8px;margin-bottom:14px}
.evi{width:36px;height:36px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0}
.evt{font-size:15px;font-weight:500}
.evs{font-size:12px;color:#6b6b6b;margin-top:1px}
.bb{background:#f5f5f0;border-radius:8px;padding:8px 12px;margin-bottom:10px;display:flex;align-items:center;justify-content:space-between}
.gr{margin-top:10px;padding:10px 12px;border-radius:8px;font-size:13px;font-weight:500}
.gn{background:#f5f5f0;color:#9b9b9b}
.gs{color:#fff}
.rg{display:flex;flex-wrap:wrap;gap:5px}
.rc{width:43px;height:43px;border-radius:8px;border:0.5px solid #d4d4d0;display:flex;flex-direction:column;align-items:center;justify-content:center;cursor:pointer;gap:1px}
.rcn{font-size:13px;font-weight:500}
.rcm{font-size:9px;color:#6b6b6b;max-width:39px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.rc.done{background:#EAF3DE;border-color:#9FE1CB}.rc.done .rcn{color:#3B6D11}
.rc.part{background:#FAEEDA;border-color:#FAC775}.rc.part .rcn{color:#854F0B}
.rc.abs{background:#FCEBEB;border-color:#F0997B}.rc.abs .rcn{color:#A32D2D}
.rc.cur{border:2px solid #378ADD;background:#E6F1FB}.rc.cur .rcn{color:#185FA5}
.sbr{background:#fff;border:0.5px solid #e8e8e4;border-radius:12px;padding:12px 14px;display:flex;align-items:center;gap:12px;margin-bottom:8px}
.sbc{font-size:14px;font-weight:500;min-width:58px;line-height:1.4}
.svb{position:sticky;bottom:0;background:#fff;border-top:0.5px solid #e8e8e4;padding:10px 12px}
.done-hd{background:#fff;border:0.5px solid #e8e8e4;border-radius:12px;padding:16px;margin-bottom:10px;text-align:center}
.done-stats{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:10px}
.dsc{background:#f5f5f0;border-radius:8px;padding:10px 6px;text-align:center}
.dsl{font-size:11px;color:#6b6b6b;margin-bottom:3px}
.dsv{font-size:20px;font-weight:500}
.ds{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:10px}
.sc{background:#f5f5f0;border-radius:8px;padding:10px 12px}
.sl{font-size:11px;color:#6b6b6b;margin-bottom:2px}
.sv{font-size:22px;font-weight:500}
table{width:100%;border-collapse:collapse;font-size:12px;table-layout:fixed}
thead th{background:#f5f5f0;padding:6px 5px;text-align:left;font-weight:500;color:#6b6b6b;font-size:11px;border-bottom:0.5px solid #e8e8e4}
tbody tr{border-bottom:0.5px solid #e8e8e4}
tbody tr:last-child{border-bottom:none}
tbody td{padding:6px 5px;vertical-align:middle}
tbody tr:hover{background:#f5f5f0}
.bdg{display:inline-block;font-size:10px;padding:2px 5px;border-radius:8px;font-weight:500}
.bd{background:#EAF3DE;color:#3B6D11}.bp2{background:#FAEEDA;color:#854F0B}
.ba{background:#FCEBEB;color:#A32D2D}.bn{background:#F1EFE8;color:#5F5E5A}
.ibt{width:25px;height:25px;border:0.5px solid #e8e8e4;border-radius:8px;background:transparent;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;color:#6b6b6b;font-size:12px}
.ibt:hover{background:#f5f5f0}
.gp{display:inline-block;width:20px;height:20px;border-radius:50%;font-size:11px;font-weight:500;line-height:20px;text-align:center}
.gp1{background:#EAF3DE;color:#3B6D11}.gp2{background:#E6F1FB;color:#185FA5}
.gp3{background:#FAEEDA;color:#854F0B}.gp4{background:#FAECE7;color:#993C1D}
.gp5{background:#FCEBEC;color:#A32D2D}.gpX{background:#f5f5f0;color:#9b9b9b}
.mb{position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:50;display:none;align-items:flex-end;justify-content:center}
.mb.open{display:flex}
.md{background:#fff;border-radius:16px 16px 0 0;padding:20px 16px 36px;width:100%;max-width:420px;border-top:0.5px solid #e8e8e4;max-height:85vh;overflow-y:auto}
.mh{font-size:15px;font-weight:500;margin-bottom:14px;display:flex;justify-content:space-between;align-items:center}
.csv-preview{background:#f5f5f0;border-radius:8px;padding:10px;font-size:11px;font-family:monospace;white-space:pre;overflow-x:auto;max-height:200px;overflow-y:auto;margin-bottom:12px;border:0.5px solid #e8e8e4;color:#6b6b6b}
.pin-overlay{position:fixed;inset:0;background:#fff;z-index:100;display:none;flex-direction:column;align-items:center;justify-content:center;padding:32px 24px}
.pin-overlay.show{display:flex}
.pin-dots{display:flex;gap:8px;margin-bottom:24px;justify-content:center;flex-wrap:wrap;max-width:280px}
.pin-dot{width:14px;height:14px;border-radius:50%;background:#ebebeb;border:1.5px solid #d4d4d0;transition:all .15s}
.pin-dot.filled{background:#378ADD;border-color:#378ADD}
.pin-dots.err .pin-dot.filled{background:#D85A30;border-color:#D85A30}
.pin-keypad{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;width:100%;max-width:280px}
.pk{height:58px;border:0.5px solid #d4d4d0;border-radius:12px;background:#fff;font-size:22px;font-weight:500;cursor:pointer;display:flex;align-items:center;justify-content:center;user-select:none}
.pk:active{background:#f5f5f0;transform:scale(.95)}
.pk.del{font-size:16px;color:#6b6b6b}
.pk.empty{background:transparent;border:none;cursor:default}
.reset-body{text-align:center;padding:8px 0 4px}
.reset-icon{font-size:40px;margin-bottom:12px}
.reset-title{font-size:16px;font-weight:500;margin-bottom:8px}
.reset-sub{font-size:13px;color:#6b6b6b;line-height:1.6;margin-bottom:24px}
.info-banner{background:#f5f5f0;border-radius:12px;padding:12px 14px;margin-bottom:12px;display:flex;align-items:flex-start;gap:10px;border:0.5px solid #e8e8e4}
.toast{position:fixed;bottom:80px;left:50%;transform:translateX(-50%);padding:8px 16px;border-radius:8px;font-size:13px;font-weight:500;display:none;z-index:300;white-space:nowrap;color:#fff}
</style>
</head>
<body>

<div class="pin-overlay" id="pin_overlay">
  <div style="font-size:48px;margin-bottom:16px">🔐</div>
  <div style="font-size:18px;font-weight:500;margin-bottom:6px">관리자 인증</div>
  <div style="font-size:13px;color:#6b6b6b;margin-bottom:28px" id="pin_sub_txt">PIN을 입력하세요 (4~8자리)</div>
  <div class="pin-dots" id="pin_dots">
    <div class="pin-dot" id="d0"></div><div class="pin-dot" id="d1"></div>
    <div class="pin-dot" id="d2"></div><div class="pin-dot" id="d3"></div>
    <div class="pin-dot" id="d4"></div><div class="pin-dot" id="d5"></div>
    <div class="pin-dot" id="d6"></div><div class="pin-dot" id="d7"></div>
  </div>
  <div class="pin-keypad">
    <button class="pk" onclick="pk('1')">1</button><button class="pk" onclick="pk('2')">2</button><button class="pk" onclick="pk('3')">3</button>
    <button class="pk" onclick="pk('4')">4</button><button class="pk" onclick="pk('5')">5</button><button class="pk" onclick="pk('6')">6</button>
    <button class="pk" onclick="pk('7')">7</button><button class="pk" onclick="pk('8')">8</button><button class="pk" onclick="pk('9')">9</button>
    <button class="pk empty"></button><button class="pk" onclick="pk('0')">0</button><button class="pk del" onclick="pkDel()">⌫</button>
  </div>
  <button style="margin-top:24px;background:transparent;border:none;cursor:pointer;font-size:13px;color:#9b9b9b" onclick="pinCancel()">취소</button>
</div>

<div class="app">
<div class="topbar">
  <button class="tb-back" id="tb_back" onclick="goBack()"><i class="ti ti-arrow-left"></i></button>
  <span class="tb-title" id="tb_title">PAPS 측정</span>
  <button class="admin-btn" id="admin_btn" onclick="adminLogout()" title="클릭하면 로그아웃">
    🛡️ 관리자 <span style="opacity:.7;font-size:10px">✕</span>
  </button>
  <button class="tb-act" id="tb_act" onclick="requireAdmin(()=>{rDash();showScreen('dash');})">대시보드</button>
</div>

<div class="screen on" id="sc_home">
  <div class="tab-bar" id="ht_tabs">
    <button class="tab on" onclick="homeTab(0)">반 선택</button>
    <button class="tab" onclick="homeTab(1)">명단 관리</button>
    <button class="tab" onclick="homeTab(2)">현황판</button>
  </div>
  <div id="ht0">
    <div class="card"><div class="ct"><i class="ti ti-calendar"></i> 기본 설정</div>
      <div class="g2"><div class="fg" style="margin-bottom:0"><label>측정일</label><input type="date" id="g_date"></div><div class="fg" style="margin-bottom:0"><label>측정자</label><input type="text" id="g_msr" placeholder="3학년 1반 반장"></div></div>
    </div>
    <div class="card"><div class="ct"><i class="ti ti-door-enter"></i> 반 선택 및 종목</div>
      <div class="g3" style="margin-bottom:10px">
        <div class="fg" style="margin-bottom:0"><label>학년</label><select id="s_grade" onchange="uCS('s_cls','s_grade')"><option value="">—</option><option>1</option><option>2</option><option>3</option></select></div>
        <div class="fg" style="margin-bottom:0"><label>반</label><select id="s_cls"><option value="">—</option></select></div>
        <div class="fg" style="margin-bottom:0"><label>측정 종목</label>
          <select id="s_event"><option value="cardio">왕복오래달리기</option><option value="grip">악력</option><option value="jump">제자리멀리뛰기</option><option value="bmi">체질량지수</option><option value="flex">앉아윗몸앞으로굽히기</option></select>
        </div>
      </div>
      <button class="bp" onclick="startClass()">이 반 측정 시작 →</button>
    </div>
    <div id="rw" style="display:none">
      <div style="font-size:12px;font-weight:500;color:#6b6b6b;margin-bottom:8px"><i class="ti ti-history"></i> 진행 중인 반</div>
      <div id="rl"></div>
    </div>
  </div>
  <div id="ht1" style="display:none">
    <div class="card">
      <div class="ct"><i class="ti ti-file-spreadsheet"></i> 엑셀 파일로 명단 불러오기</div>
      <div class="dropzone" id="dropzone" onclick="document.getElementById('xlsx_input').click()"
        ondragover="event.preventDefault();this.classList.add('drag')"
        ondragleave="this.classList.remove('drag')"
        ondrop="handleDrop(event)">
        <div style="font-size:28px;margin-bottom:8px">📊</div>
        <div style="font-size:13px;font-weight:500;margin-bottom:4px">엑셀 파일 클릭 또는 드래그</div>
        <div style="font-size:12px;color:#6b6b6b">학년 / 반명 / 번호 / 학생성명 포함 양식</div>
      </div>
      <input type="file" id="xlsx_input" accept=".xlsx,.xls" onchange="handleXlsxFile(this.files[0])">
      <div class="ur" id="ur"></div>
      <div id="xp" style="display:none">
        <div style="font-size:12px;font-weight:500;color:#6b6b6b;margin-bottom:8px">불러온 반 목록</div>
        <div id="xcl"></div>
        <button class="bp" onclick="saveAllXlsx()" style="margin-top:8px">전체 반 저장</button>
      </div>
    </div>
    <div class="card"><div class="ct"><i class="ti ti-keyboard"></i> 직접 입력</div>
      <div class="g2" style="margin-bottom:8px">
        <div class="fg" style="margin-bottom:0"><label>학년</label><select id="r_grade" onchange="uCS('r_cls','r_grade')"><option value="">—</option><option>1</option><option>2</option><option>3</option></select></div>
        <div class="fg" style="margin-bottom:0"><label>반</label><select id="r_cls"><option value="">—</option></select></div>
      </div>
      <div class="fg"><label>이름 목록 (번호,이름,성별 또는 이름만)</label>
        <textarea id="r_txt" rows="5" placeholder="1,홍길동,M&#10;2,김영희,F&#10;..."></textarea>
      </div>
      <button class="bp" onclick="importR()" style="margin-bottom:8px">명단 저장</button>
      <button class="bs" onclick="showRP()">저장된 명단 보기</button>
    </div>
  </div>
  <div id="ht2" style="display:none">
    <div class="info-banner">
      <span style="font-size:18px;flex-shrink:0">📡</span>
      <div>
        <div style="font-size:13px;font-weight:500;margin-bottom:2px">실시간 측정 현황</div>
        <div style="font-size:12px;color:#6b6b6b">반장이 입력하면 여기에 반영됩니다. 새로고침 버튼으로 최신 상태를 확인하세요.</div>
      </div>
    </div>
    <div class="card">
      <div class="ct" style="justify-content:space-between;margin-bottom:6px">
        <span><i class="ti ti-layout-board"></i> 반별 측정 현황</span>
        <button class="refresh-btn" onclick="doRefresh()"><i class="ti ti-refresh" id="rbtn_icon"></i></button>
      </div>
      <div id="sb_ts" style="font-size:11px;color:#9b9b9b;margin-bottom:10px;text-align:right"></div>
      <div id="status_board"><div style="text-align:center;padding:20px;color:#9b9b9b;font-size:13px">측정을 시작하면 여기에 표시됩니다</div></div>
    </div>
  </div>
</div>

<div class="screen" id="sc_roster">
  <div class="card" style="margin-bottom:8px"><div class="ct"><i class="ti ti-layout-grid"></i> 번호 현황 — 클릭하면 바로 이동</div><div class="rg" id="rc_grid"></div></div>
  <div class="g2"><button class="bs" onclick="goHome()">← 반 목록</button><button class="bp" onclick="startFF()">순서대로 시작 →</button></div>
</div>

<div class="screen" id="sc_input" style="padding-bottom:80px">
  <div class="sn">
    <div class="sti">
      <div class="sid">
        <div class="sci" id="i_circle">1</div>
        <div><div class="snm" id="i_nd">이름</div><div class="ssb" id="i_sub">1학년 1반</div></div>
      </div>
      <div style="display:flex;gap:5px">
        <button class="bsm" style="color:#D85A30;border-color:#F0997B" onclick="markAbs()">결석</button>
        <button class="bsm" onclick="skipS()">건너뜀</button>
      </div>
    </div>
    <div class="pb-box">
      <div>
        <div style="display:flex;align-items:baseline;gap:6px">
          <span style="font-size:26px;font-weight:500" id="i_cn">1</span>
          <span style="font-size:13px;color:#6b6b6b">번째 측정 중</span>
        </div>
        <div style="font-size:13px;color:#6b6b6b;margin-top:3px">
          <span style="font-weight:500;color:#1D9E75" id="i_done_num">0</span>명 완료
        </div>
      </div>
      <div style="text-align:right">
        <div style="font-size:13px;color:#6b6b6b">전체 <span style="font-weight:500" id="i_tn">30</span>명</div>
        <div style="font-size:13px;font-weight:500;color:#378ADD;margin-top:2px" id="i_pt">0%</div>
      </div>
    </div>
    <div style="margin-bottom:10px">
      <div class="pbar"><div class="pbf" id="i_pb" style="width:0%"></div></div>
      <div class="pbls"><span>1번</span><span id="i_tl">30번</span></div>
    </div>
    <div class="nr">
      <button class="nb" onclick="prevS()"><i class="ti ti-chevron-left"></i> 이전</button>
      <button class="nb" id="i_nb" onclick="nextS()">다음 <i class="ti ti-chevron-right"></i></button>
    </div>
  </div>
  <div class="card" style="margin-bottom:8px">
    <div class="g2">
      <div class="fg" style="margin-bottom:0"><label>이름</label><input type="text" id="i_name" placeholder="홍길동"></div>
      <div class="fg" style="margin-bottom:0"><label>성별</label><select id="i_gen"><option value="">—</option><option value="M">남</option><option value="F">여</option></select></div>
    </div>
  </div>
  <div class="card" style="margin-bottom:0">
    <div class="etb" id="ev_tabs"></div>
    <div id="p_cardio" class="evb" style="display:none">
      <div class="evh"><div class="evi" style="background:#E6F1FB"><i class="ti ti-run" style="color:#378ADD;font-size:20px"></i></div><div><div class="evt">왕복오래달리기</div><div class="evs">심폐지구력 · 20m</div></div></div>
      <div class="g2"><div class="fg" style="margin-bottom:0"><label>완료 횟수 (회)</label><input type="number" id="i_cardio" placeholder="0" min="0" oninput="ag('cardio')"></div><div class="fg" style="margin-bottom:0"><label>등급 (자동)</label><input type="number" id="i_cardioG" readonly placeholder="—"></div></div>
      <div class="gr gn" id="gr_cardio">횟수를 입력하면 자동 계산됩니다</div>
    </div>
    <div id="p_grip" class="evb" style="display:none">
      <div class="evh"><div class="evi" style="background:#E1F5EE"><i class="ti ti-hand-grab" style="color:#1D9E75;font-size:20px"></i></div><div><div class="evt">악력</div><div class="evs">우1→좌1→우2→좌2 순서</div></div></div>
      <div style="font-size:11px;color:#6b6b6b;margin-bottom:10px;padding:6px 10px;background:#f5f5f0;border-radius:8px">좌·우 각각 높은 값의 평균으로 등급 산출</div>
      <div class="g4" style="margin-bottom:8px">
        <div class="fg" style="margin-bottom:0"><label>우 1회</label><input type="number" id="i_gr1" placeholder="0.0" step="0.1" oninput="agGrip()"></div>
        <div class="fg" style="margin-bottom:0"><label>좌 1회</label><input type="number" id="i_gl1" placeholder="0.0" step="0.1" oninput="agGrip()"></div>
        <div class="fg" style="margin-bottom:0"><label>우 2회</label><input type="number" id="i_gr2" placeholder="0.0" step="0.1" oninput="agGrip()"></div>
        <div class="fg" style="margin-bottom:0"><label>좌 2회</label><input type="number" id="i_gl2" placeholder="0.0" step="0.1" oninput="agGrip()"></div>
      </div>
      <div class="bb" id="grip_bb" style="display:none"><div style="font-size:12px;color:#6b6b6b">우 최고 <span id="gbr" style="font-weight:500">—</span>kg | 좌 최고 <span id="gbl" style="font-weight:500">—</span>kg | 평균 <span id="gav" style="font-weight:500">—</span>kg</div></div>
      <div class="g2" style="margin-top:8px"><div class="fg" style="margin-bottom:0"><label>평균 (자동)</label><input type="number" id="i_grip" readonly placeholder="—"></div><div class="fg" style="margin-bottom:0"><label>등급 (자동)</label><input type="number" id="i_gripG" readonly placeholder="—"></div></div>
      <div class="gr gn" id="gr_grip">측정값을 입력하면 자동 계산됩니다</div>
    </div>
    <div id="p_jump" class="evb" style="display:none">
      <div class="evh"><div class="evi" style="background:#FAEEDA"><i class="ti ti-arrows-up" style="color:#BA7517;font-size:20px"></i></div><div><div class="evt">제자리멀리뛰기</div><div class="evs">2회 측정, 좋은 기록 채택</div></div></div>
      <div class="g2" style="margin-bottom:8px">
        <div class="fg" style="margin-bottom:0"><label>1회 (cm)</label><input type="number" id="i_j1" placeholder="0.0" step="0.1" oninput="agJump()"></div>
        <div class="fg" style="margin-bottom:0"><label>2회 (cm)</label><input type="number" id="i_j2" placeholder="0.0" step="0.1" oninput="agJump()"></div>
      </div>
      <div class="bb" id="jump_bb" style="display:none"><div style="font-size:12px;color:#6b6b6b">채택</div><div style="font-size:15px;font-weight:500" id="jb">—</div></div>
      <div class="g2" style="margin-top:8px"><div class="fg" style="margin-bottom:0"><label>채택 (자동)</label><input type="number" id="i_jump" readonly placeholder="—"></div><div class="fg" style="margin-bottom:0"><label>등급 (자동)</label><input type="number" id="i_jumpG" readonly placeholder="—"></div></div>
      <div class="gr gn" id="gr_jump">측정값을 입력하면 자동 계산됩니다</div>
    </div>
    <div id="p_bmi" class="evb" style="display:none">
      <div class="evh"><div class="evi" style="background:#FBEAF0"><i class="ti ti-ruler-measure" style="color:#993556;font-size:20px"></i></div><div><div class="evt">체질량지수 (BMI)</div><div class="evs">키와 몸무게 측정</div></div></div>
      <div class="g2"><div class="fg" style="margin-bottom:0"><label>키 (cm)</label><input type="number" id="i_ht" placeholder="170" step="0.1" oninput="ag('bmi')"></div><div class="fg" style="margin-bottom:0"><label>몸무게 (kg)</label><input type="number" id="i_wt" placeholder="60" step="0.1" oninput="ag('bmi')"></div></div>
      <div class="gr gn" id="gr_bmi">키와 몸무게를 입력하면 자동 계산됩니다</div>
    </div>
    <div id="p_flex" class="evb" style="display:none">
      <div class="evh"><div class="evi" style="background:#EEEDFE"><i class="ti ti-body-scan" style="color:#534AB7;font-size:20px"></i></div><div><div class="evt">앉아윗몸앞으로굽히기</div><div class="evs">2회 측정, 좋은 기록 채택</div></div></div>
      <div class="g2" style="margin-bottom:8px">
        <div class="fg" style="margin-bottom:0"><label>1회 (cm)</label><input type="number" id="i_f1" placeholder="0.0" step="0.1" oninput="agFlex()"></div>
        <div class="fg" style="margin-bottom:0"><label>2회 (cm)</label><input type="number" id="i_f2" placeholder="0.0" step="0.1" oninput="agFlex()"></div>
      </div>
      <div class="bb" id="flex_bb" style="display:none"><div style="font-size:12px;color:#6b6b6b">채택</div><div style="font-size:15px;font-weight:500" id="fb">—</div></div>
      <div class="g2" style="margin-top:8px"><div class="fg" style="margin-bottom:0"><label>채택 (자동)</label><input type="number" id="i_flex" readonly placeholder="—"></div><div class="fg" style="margin-bottom:0"><label>등급 (자동)</label><input type="number" id="i_flexG" readonly placeholder="—"></div></div>
      <div class="gr gn" id="gr_flex">측정값을 입력하면 자동 계산됩니다</div>
    </div>
    <div style="border-top:0.5px solid #e8e8e4;margin-top:14px;padding-top:12px">
      <div class="fg" style="margin-bottom:0"><label>메모</label><input type="text" id="i_memo" placeholder="부상, 특이사항 등"></div>
    </div>
  </div>
  <div class="svb"><div class="g2"><button class="bs" style="height:44px" onclick="showScreen('roster')"><i class="ti ti-layout-grid"></i> 목록</button><button class="bp" onclick="saveNext()"><i class="ti ti-device-floppy"></i> 저장 후 다음</button></div></div>
</div>

<div class="screen" id="sc_done">
  <div class="done-hd">
    <div style="font-size:28px;margin-bottom:8px">✅</div>
    <div style="font-size:17px;font-weight:500;margin-bottom:4px" id="done_title">측정 완료</div>
    <div style="font-size:13px;color:#6b6b6b" id="done_sub">기록을 확인하세요</div>
  </div>
  <div class="done-stats" id="done_stats"></div>
  <div class="card" style="margin-bottom:10px">
    <div class="ct" id="done_ct"><i class="ti ti-list-check"></i> 결과</div>
    <div style="overflow-x:auto"><table><thead id="done_thead"></thead><tbody id="done_body"></tbody></table></div>
  </div>
  <div class="g2">
    <button class="bs" onclick="showScreen('roster')">번호 목록</button>
    <button class="bp" onclick="goHomeTab2()">현황판 보기 →</button>
  </div>
</div>

<div class="screen" id="sc_dash">
  <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;flex-wrap:wrap">
    <span style="font-size:14px;font-weight:500">📊 교사 대시보드</span>
    <span style="font-size:11px;padding:2px 8px;border-radius:20px;background:#EAF3DE;color:#3B6D11;font-weight:500">🛡️ 관리자 전용</span>
  </div>
  <div class="ds" id="d_stats"></div>
  <div class="g2" style="margin-bottom:8px">
    <select id="df_g" onchange="rDash()"><option value="">전체 학년</option><option>1학년</option><option>2학년</option><option>3학년</option></select>
    <select id="df_s" onchange="rDash()"><option value="">전체 상태</option><option value="done">완료</option><option value="partial">일부</option><option value="absent">결석</option><option value="none">미입력</option></select>
  </div>
  <div style="margin-bottom:8px"><input type="text" id="df_q" placeholder="이름·반 검색..." oninput="rDash()" style="width:100%"></div>
  <div style="border:0.5px solid #e8e8e4;border-radius:12px;overflow:hidden;margin-bottom:10px">
    <table><thead><tr>
      <th style="width:50px">학년반번</th><th style="width:44px">이름</th>
      <th style="width:28px">심폐</th><th style="width:28px">악력</th><th style="width:28px">순발</th><th style="width:28px">BMI</th><th style="width:28px">유연</th>
      <th style="width:42px">상태</th><th style="width:42px">관리</th>
    </tr></thead><tbody id="d_body"></tbody></table>
  </div>
  <div id="d_empty" style="display:none;text-align:center;padding:24px;color:#9b9b9b;font-size:13px">기록 없음</div>
  <div class="g2" style="margin-bottom:8px">
    <button class="bs" onclick="showCsvModal()"><i class="ti ti-download"></i> CSV 내보내기</button>
    <button class="bs" style="color:#D85A30;border-color:#F0997B" onclick="openResetModal()"><i class="ti ti-trash"></i> 초기화</button>
  </div>
  <button class="bs" style="width:100%;margin-bottom:8px" onclick="openPinChangeModal()"><i class="ti ti-key"></i> 관리자 PIN 변경</button>
  <button class="bs" style="width:100%;color:#888780" onclick="adminLogout();showScreen('home')"><i class="ti ti-logout"></i> 관리자 로그아웃</button>
</div>
</div>

<div class="toast" id="toast"></div>
<div class="mb" id="em"><div class="md"><div class="mh">기록 수정 <span style="cursor:pointer;color:#6b6b6b;font-size:22px;line-height:1" onclick="closeM('em')">×</span></div><div id="em_b"></div></div></div>
<div class="mb" id="rm"><div class="md"><div class="mh">저장된 명단 <span style="cursor:pointer;color:#6b6b6b;font-size:22px;line-height:1" onclick="closeM('rm')">×</span></div><div id="rm_b"></div></div></div>
<div class="mb" id="csv_modal">
  <div class="md">
    <div class="mh">CSV 내보내기 <span style="cursor:pointer;color:#6b6b6b;font-size:22px;line-height:1" onclick="closeM('csv_modal')">×</span></div>
    <p style="font-size:13px;color:#6b6b6b;margin-bottom:10px">복사 후 메모장 붙여넣기 → <strong>.csv</strong>로 저장</p>
    <div class="csv-preview" id="csv_content"></div>
    <div class="g2"><button class="bs" onclick="copyCsv()"><i class="ti ti-copy"></i> 클립보드 복사</button><button class="bp" onclick="closeM('csv_modal')">닫기</button></div>
    <div id="copy_ok" style="display:none;text-align:center;margin-top:8px;font-size:13px;color:#1D9E75;font-weight:500">✓ 클립보드에 복사되었습니다!</div>
  </div>
</div>
<div class="mb" id="reset_modal">
  <div class="md">
    <div class="reset-body">
      <div class="reset-icon">⚠️</div>
      <div class="reset-title">모든 기록을 삭제할까요?</div>
      <div class="reset-sub">저장된 측정 기록이 모두 삭제됩니다.<br>이 작업은 되돌릴 수 없습니다.</div>
      <div class="g2" style="gap:10px">
        <button class="bs" style="height:44px" onclick="closeM('reset_modal')">취소</button>
        <button class="bs" style="height:44px;color:#D85A30;border-color:#F0997B;font-weight:500" onclick="doReset()">삭제 확인</button>
      </div>
    </div>
  </div>
</div>
<div class="mb" id="pin_change_modal">
  <div class="md">
    <div class="mh">관리자 PIN 변경 <span style="cursor:pointer;color:#6b6b6b;font-size:22px;line-height:1" onclick="closeM('pin_change_modal')">×</span></div>
    <div class="fg"><label>새 PIN (4~8자리 숫자)</label><input type="password" id="pin_new1" placeholder="새 PIN" maxlength="8" inputmode="numeric"></div>
    <div class="fg"><label>새 PIN 확인</label><input type="password" id="pin_new2" placeholder="새 PIN 재입력" maxlength="8" inputmode="numeric"></div>
    <div id="pin_change_err" style="font-size:12px;color:#D85A30;margin-bottom:8px;display:none"></div>
    <button class="bp" onclick="savePinChange()">변경 저장</button>
  </div>
</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script>
const SK='paps_v7',PK='paps_pin',DEFAULT_PIN='1234';
let records=[],rosters={},classes=[];
let cc=null,ci=0,ei=null,cet='cardio';
let xlsxParsed={},csvText='';
let adminUnlocked=false,pinBuf='',pinCallback=null;

function getPin(){return localStorage.getItem(PK)||DEFAULT_PIN;}
function setPin(p){localStorage.setItem(PK,p);}
function pk(d){if(pinBuf.length>=8)return;pinBuf+=d;renderDots();if(pinBuf.length===4)setTimeout(checkPin,100);}
function pkDel(){pinBuf=pinBuf.slice(0,-1);renderDots();}
function renderDots(){for(let i=0;i<8;i++){const el=document.getElementById('d'+i);if(el)el.classList.toggle('filled',i<pinBuf.length);}document.getElementById('pin_dots').classList.remove('err');}
function checkPin(){if(pinBuf===getPin()){document.getElementById('pin_overlay').classList.remove('show');adminUnlocked=true;document.getElementById('admin_btn').style.display='flex';pinBuf='';renderDots();if(pinCallback){const cb=pinCallback;pinCallback=null;cb();}}else{document.getElementById('pin_dots').classList.add('err');setTimeout(()=>{pinBuf='';renderDots();},600);}}
function pinCancel(){document.getElementById('pin_overlay').classList.remove('show');pinBuf='';renderDots();pinCallback=null;}
function requireAdmin(cb){if(adminUnlocked){cb();return;}pinCallback=cb;document.getElementById('pin_overlay').classList.add('show');}
function adminLogout(){adminUnlocked=false;document.getElementById('admin_btn').style.display='none';showToast('관리자 로그아웃','#888780');}
function openPinChangeModal(){document.getElementById('pin_new1').value='';document.getElementById('pin_new2').value='';document.getElementById('pin_change_err').style.display='none';document.getElementById('pin_change_modal').classList.add('open');}
function savePinChange(){const n1=document.getElementById('pin_new1').value.trim(),n2=document.getElementById('pin_new2').value.trim();const err=document.getElementById('pin_change_err');if(!/^\d{4,8}$/.test(n1)){err.textContent='4~8자리 숫자로 입력해주세요';err.style.display='block';return;}if(n1!==n2){err.textContent='새 PIN이 일치하지 않습니다';err.style.display='block';return;}setPin(n1);closeM('pin_change_modal');showToast('✓ PIN이 변경되었습니다');}
const EV={cardio:{l:'심폐',f:'왕복오래달리기',c:'#378ADD'},grip:{l:'악력',f:'악력',c:'#1D9E75'},jump:{l:'순발',f:'제자리멀리뛰기',c:'#BA7517'},bmi:{l:'BMI',f:'체질량지수',c:'#993556'},flex:{l:'유연',f:'앉아윗몸앞으로굽히기',c:'#534AB7'}};
const EK=Object.keys(EV);
const CARDIO_B={M:{1:[25,26,41,42,55,56,69,70],2:[27,28,43,44,57,58,71,72],3:[29,30,45,46,59,60,73,74]},F:{1:[16,17,24,25,36,37,49,50],2:[17,18,26,27,40,41,54,55],3:[17,18,26,27,40,41,54,55]}};
const GRIP_B={M:{1:[26,27,33,34,40,41,47,48],2:[26,27,33,34,40,41,47,48],3:[26,27,33,34,40,41,47,48]},F:{1:[15,16,20,21,25,26,30,31],2:[15,16,20,21,25,26,30,31],3:[15,16,20,21,25,26,30,31]}};
const JUMP_B={M:{1:[149,150,174,175,199,200,224,225],2:[149,150,174,175,199,200,224,225],3:[149,150,174,175,199,200,224,225]},F:{1:[109,110,129,130,149,150,169,170],2:[109,110,129,130,149,150,169,170],3:[109,110,129,130,149,150,169,170]}};
const FLEX_B={M:{1:[4,5,11,12,18,19,25,26],2:[4,5,11,12,18,19,25,26],3:[4,5,11,12,18,19,25,26]},F:{1:[8,9,15,16,22,23,29,30],2:[8,9,15,16,22,23,29,30],3:[8,9,15,16,22,23,29,30]}};
function grd(v,b){if(v<=b[0])return 5;if(v<=b[2])return 4;if(v<=b[4])return 3;if(v<=b[6])return 2;return 1;}
function cG(){return document.getElementById('i_gen').value||(cc?cc.genders[ci]||'':'');}
function cGr(){return cc?cc.grade:'1';}
const GL=['','매우 좋음','좋음','보통','나쁨','매우 나쁨'];
function setGR(ev,gr,extra){const el=document.getElementById('gr_'+ev);if(!el)return;if(!gr){el.className='gr gn';el.style.background='';el.textContent='측정값을 입력하면 자동 계산됩니다';return;}el.className='gr gs';el.style.background=EV[ev].c;el.textContent=(extra?extra+' · ':'')+gr+'등급  '+GL[gr];}
function ag(ev){const g=cG(),gr=cGr();if(ev==='cardio'){const v=parseFloat(document.getElementById('i_cardio').value);if(isNaN(v))return;const b=CARDIO_B[g]?.[gr]||CARDIO_B[g]?.['1'];const grade=g&&b?grd(v,b):'';document.getElementById('i_cardioG').value=grade;setGR('cardio',grade);}else if(ev==='bmi'){const h=parseFloat(document.getElementById('i_ht').value),w=parseFloat(document.getElementById('i_wt').value);if(h>0&&w>0){const b=Math.round(w/((h/100)**2)*10)/10;const grade=b<18.5?3:b<23?1:b<25?2:b<30?4:5;const lbl=['','정상','과체중','저체중','비만','고도비만'];setGR('bmi',grade,'BMI '+b+' ('+lbl[grade]+')');}}}
function agGrip(){const r1=parseFloat(document.getElementById('i_gr1').value)||0,l1=parseFloat(document.getElementById('i_gl1').value)||0;const r2=parseFloat(document.getElementById('i_gr2').value)||0,l2=parseFloat(document.getElementById('i_gl2').value)||0;if(!r1&&!l1&&!r2&&!l2)return;const bR=Math.max(r1,r2),bL=Math.max(l1,l2),avg=bR&&bL?Math.round((bR+bL)/2*10)/10:bR||bL;document.getElementById('grip_bb').style.display='block';document.getElementById('gbr').textContent=bR||'—';document.getElementById('gbl').textContent=bL||'—';document.getElementById('gav').textContent=avg||'—';document.getElementById('i_grip').value=avg||'';const g=cG(),gr=cGr(),b=GRIP_B[g]?.[gr]||GRIP_B[g]?.['1'];const grade=avg&&g&&b?grd(avg,b):'';document.getElementById('i_gripG').value=grade;setGR('grip',grade,'좌우 평균 '+avg+'kg');}
function agJump(){const v1=parseFloat(document.getElementById('i_j1').value),v2=parseFloat(document.getElementById('i_j2').value);const best=Math.max(isNaN(v1)?0:v1,isNaN(v2)?0:v2);if(!best)return;document.getElementById('jump_bb').style.display='block';document.getElementById('jb').textContent=best+'cm';document.getElementById('i_jump').value=best;const g=cG(),gr=cGr(),b=JUMP_B[g]?.[gr]||JUMP_B[g]?.['1'];const grade=g&&b?grd(best,b):'';document.getElementById('i_jumpG').value=grade;setGR('jump',grade,'채택 '+best+'cm');}
function agFlex(){const v1=parseFloat(document.getElementById('i_f1').value),v2=parseFloat(document.getElementById('i_f2').value);const vld=[v1,v2].filter(v=>!isNaN(v));if(!vld.length)return;const best=Math.max(...vld);document.getElementById('flex_bb').style.display='block';document.getElementById('fb').textContent=best+'cm';document.getElementById('i_flex').value=best;const g=cG(),gr=cGr(),b=FLEX_B[g]?.[gr]||FLEX_B[g]?.['1'];const grade=g&&b?grd(best,b):'';document.getElementById('i_flexG').value=grade;setGR('flex',grade,'채택 '+best+'cm');}
document.getElementById('i_gen').addEventListener('change',()=>{ag('cardio');agGrip();agJump();ag('bmi');agFlex();});
function handleDrop(e){e.preventDefault();document.getElementById('dropzone').classList.remove('drag');const f=e.dataTransfer.files[0];if(f)handleXlsxFile(f);}
function handleXlsxFile(file){if(!file)return;if(!file.name.match(/\.(xlsx|xls)$/i)){showUR('err','xlsx/xls 파일만 지원합니다');return;}const rd=new FileReader();rd.onload=e=>{try{const wb=XLSX.read(new Uint8Array(e.target.result),{type:'array'});parseRows(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:''}),file.name);}catch(err){showUR('err','파일 읽기 오류: '+err.message);}};rd.readAsArrayBuffer(file);}
function parseRows(rows,fname){let hi=-1;for(let i=0;i<Math.min(rows.length,5);i++){const r=rows[i].map(c=>String(c||''));if(r.some(c=>c.includes('번호'))&&r.some(c=>c.includes('성명')||c.includes('이름'))){hi=i;break;}}if(hi<0){showUR('err','헤더를 찾을 수 없습니다');return;}const hd=rows[hi].map(c=>String(c||'').trim());const iG=hd.findIndex(h=>h==='학년'),iC=hd.findIndex(h=>h==='반명'||h==='반');const iN=hd.findIndex(h=>h==='번호'),iNm=hd.findIndex(h=>h.includes('성명')||h==='이름');const iSx=hd.findIndex(h=>h.includes('성별'));if(iG<0||iC<0||iN<0||iNm<0){showUR('err','학년/반명/번호/성명 열을 찾을 수 없습니다');return;}xlsxParsed={};for(let i=hi+1;i<rows.length;i++){const row=rows[i],grade=String(row[iG]||'').trim(),cls=String(row[iC]||'').trim();const num=parseInt(row[iN])||0,name=String(row[iNm]||'').trim();const sx=iSx>=0?String(row[iSx]||'').trim():'';const gender=sx==='1'||sx==='남'||sx==='M'?'M':sx==='2'||sx==='여'||sx==='F'?'F':'';if(!grade||!cls||!num||!name)continue;const key=grade+'_'+cls;if(!xlsxParsed[key])xlsxParsed[key]={grade,cls,students:[]};xlsxParsed[key].students.push({num,name,gender});}const keys=Object.keys(xlsxParsed);if(!keys.length){showUR('err','학생 데이터가 없습니다');return;}showUR('ok','✓ '+fname+' — '+keys.length+'개 반, '+keys.reduce((s,k)=>s+xlsxParsed[k].students.length,0)+'명 파싱 완료');document.getElementById('xp').style.display='block';document.getElementById('xcl').innerHTML=keys.sort().map(k=>{const c=xlsxParsed[k],already=rosters[k]?'(기존있음)':'';return '<div style="display:flex;align-items:center;justify-content:space-between;padding:8px 10px;background:#f5f5f0;border-radius:8px;margin-bottom:5px"><span style="font-size:13px;font-weight:500">'+c.grade+'학년 '+c.cls+'반 · '+c.students.length+'명 <span style="font-size:11px;color:#6b6b6b">'+already+'</span></span><button class="bsm" style="color:#378ADD;border-color:#378ADD" onclick="saveOneXlsx(\''+k+'\')">저장</button></div>';}).join('');}
function saveOneXlsx(key){const c=xlsxParsed[key];if(!c)return;c.students.sort((a,b)=>a.num-b.num);rosters[key]=c.students.map(s=>({num:s.num,name:s.name,gender:s.gender}));saveAll();showToast('✓ '+c.grade+'학년 '+c.cls+'반 '+c.students.length+'명 저장');rRecent();}
function saveAllXlsx(){Object.keys(xlsxParsed).forEach(k=>{xlsxParsed[k].students.sort((a,b)=>a.num-b.num);rosters[k]=xlsxParsed[k].students.map(s=>({num:s.num,name:s.name,gender:s.gender}));});saveAll();rRecent();showToast('✓ 전체 '+Object.keys(xlsxParsed).length+'개 반 저장 완료');}
function showUR(t,m){const el=document.getElementById('ur');el.className='ur '+(t==='ok'?'ok':'err');el.textContent=m;el.style.display='block';}
function importR(){const g=document.getElementById('r_grade').value,c=document.getElementById('r_cls').value,raw=document.getElementById('r_txt').value.trim();if(!g||!c){showToast('학년과 반을 선택해주세요','#D85A30');return;}const key=g+'_'+c,lines=raw.split('\n').map(l=>l.trim()).filter(Boolean),roster=[];lines.forEach((line,i)=>{const p=line.split(/[\t,]/);let num=i+1,name='',gender='';if(p.length>=3&&!isNaN(parseInt(p[0]))){num=parseInt(p[0]);name=p[1].trim();gender=p[2].trim().toUpperCase();}else if(p.length===2&&!isNaN(parseInt(p[0]))){num=parseInt(p[0]);name=p[1].trim();}else{name=p[0].trim();}roster.push({num,name,gender:gender==='M'||gender==='남'?'M':gender==='F'||gender==='여'?'F':''});});rosters[key]=roster;saveAll();showToast('✓ '+g+'학년 '+c+'반 명단 '+roster.length+'명 저장');document.getElementById('r_txt').value='';}
function showRP(){const g=document.getElementById('r_grade').value,c=document.getElementById('r_cls').value;if(!g||!c){showToast('학년과 반을 선택해주세요','#D85A30');return;}const ros=rosters[g+'_'+c];if(!ros||!ros.length){showToast('저장된 명단이 없습니다','#D85A30');return;}document.getElementById('rm_b').innerHTML='<p style="font-size:13px;color:#6b6b6b;margin-bottom:10px">'+g+'학년 '+c+'반 · '+ros.length+'명</p>'+ros.map(r=>'<div style="display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:0.5px solid #e8e8e4"><span style="font-size:12px;color:#6b6b6b;min-width:24px">'+r.num+'</span><span style="font-size:14px;font-weight:500">'+(r.name||'(이름없음)')+'</span><span style="font-size:12px;color:#6b6b6b;margin-left:auto">'+(r.gender==='M'?'남':r.gender==='F'?'여':'—')+'</span></div>').join('');document.getElementById('rm').classList.add('open');}
function loadAll(){try{const d=JSON.parse(localStorage.getItem(SK)||'{}');records=d.records||[];rosters=d.rosters||{};classes=d.classes||[];}catch(e){records=[];rosters={};classes=[];}}
function saveAll(){localStorage.setItem(SK,JSON.stringify({records,rosters,classes}));}
function openResetModal(){document.getElementById('reset_modal').classList.add('open');}
function doReset(){records=[];classes=[];cc=null;saveAll();closeM('reset_modal');rDash();rRecent();showToast('초기화 완료','#888780');}
function buildCsvText(){const h=['학년도','과정명','계열명','학년','학과명','반명','반코드','번호','학생성명','왕복오래달리기(회)','앉아윗몸 앞으로 굽히기(cm) 1차','앉아윗몸 앞으로 굽히기(cm) 2차','악력(kg) 1차 오른쪽','악력(kg) 1차 왼쪽','악력(kg) 2차 오른쪽','악력(kg) 2차 왼쪽','제자리멀리뛰기(cm) 1차','제자리멀리뛰기(cm) 2차','신장(cm)','체중(kg)'];const yr=new Date().getFullYear();const rows=[...records].sort((a,b)=>parseInt(a.grade)*10000+parseInt(a.cls)*100+parseInt(a.num)-(parseInt(b.grade)*10000+parseInt(b.cls)*100+parseInt(b.num))).map(r=>[yr,'주간','일반계',r.grade,'일반학과',r.cls,String(r.cls).padStart(2,'0'),r.num,r.name||'',r.cardio||'',r.flex1||'',r.flex2||'',r.gr_r1||'',r.gr_l1||'',r.gr_r2||'',r.gr_l2||'',r.jump1||'',r.jump2||'',r.height||'',r.weight||'']);return[h,...rows].map(r=>r.join('\t')).join('\n');}
function showCsvModal(){if(!records.length){showToast('저장된 기록이 없습니다','#D85A30');return;}csvText=buildCsvText();document.getElementById('csv_content').textContent=csvText;document.getElementById('copy_ok').style.display='none';document.getElementById('csv_modal').classList.add('open');}
function copyCsv(){if(navigator.clipboard&&navigator.clipboard.writeText){navigator.clipboard.writeText(csvText).then(()=>{document.getElementById('copy_ok').style.display='block';showToast('✓ 클립보드에 복사 완료');}).catch(()=>fbCopy());}else{fbCopy();}}
function fbCopy(){const el=document.getElementById('csv_content');const r=document.createRange();r.selectNodeContents(el);const s=window.getSelection();s.removeAllRanges();s.addRange(r);try{document.execCommand('copy');document.getElementById('copy_ok').style.display='block';showToast('✓ 복사 완료');}catch(e){showToast('Ctrl+A→Ctrl+C 해주세요','#D85A30');}}
function showScreen(name){document.querySelectorAll('.screen').forEach(s=>s.classList.remove('on'));document.getElementById('sc_'+name).classList.add('on');const back=document.getElementById('tb_back'),act=document.getElementById('tb_act');const titles={home:'PAPS 측정',roster:'번호 현황',input:cc?cc.grade+'학년 '+cc.cls+'반':'측정 입력',done:cc?cc.grade+'학년 '+cc.cls+'반 완료':'완료',dash:'교사 대시보드'};document.getElementById('tb_title').textContent=titles[name]||'PAPS';if(name==='home'||name==='roster'||name==='done'){back.style.display=(name==='home')?'none':'flex';act.style.display='';act.textContent='대시보드';act.onclick=()=>requireAdmin(()=>{rDash();showScreen('dash');});}else if(name==='input'){back.style.display='flex';act.style.display='';act.textContent='목록';act.onclick=()=>showScreen('roster');}else if(name==='dash'){back.style.display='flex';act.style.display='none';}}
function goBack(){const id=document.querySelector('.screen.on')?.id;if(id==='sc_input')showScreen('roster');else if(id==='sc_roster'||id==='sc_done'){cc=null;showScreen('home');rRecent();}else if(id==='sc_dash')showScreen('home');}
function goHome(){cc=null;showScreen('home');rRecent();}
function goHomeTab2(){showScreen('home');homeTab(2);}
function homeTab(i){
  if(i===1&&!adminUnlocked){
    requireAdmin(()=>{
      ['ht0','ht1','ht2'].forEach((id,j)=>document.getElementById(id).style.display=j===1?'block':'none');
      document.querySelectorAll('#ht_tabs .tab').forEach((t,j)=>t.classList.toggle('on',j===1));
    });
    return;
  }
  ['ht0','ht1','ht2'].forEach((id,j)=>document.getElementById(id).style.display=j===i?'block':'none');
  document.querySelectorAll('#ht_tabs .tab').forEach((t,j)=>t.classList.toggle('on',j===i));
  if(i===2){loadAll();rSB();}
}
function doRefresh(){const icon=document.getElementById('rbtn_icon');icon.classList.add('spinning');setTimeout(()=>icon.classList.remove('spinning'),500);loadAll();rSB();showToast('현황판 새로고침 완료');}
function uCS(selId,gradeId){const g=document.getElementById(gradeId).value,sel=document.getElementById(selId);sel.innerHTML='<option value="">—</option>';if(g)for(let i=1;i<=12;i++){const o=document.createElement('option');o.value=i;o.textContent=i+'반';sel.appendChild(o);}}
function buildET(){document.getElementById('ev_tabs').innerHTML=EK.map(k=>{const ev=EV[k],isOn=k===cet,isCur=k===cc?.curEvent;return '<button class="et'+(isOn?' on':'')+'" id="et_'+k+'" onclick="swET(\''+k+'\')" style="'+(isOn?'background:'+ev.c+';border-color:'+ev.c+';color:#fff':isCur?'border-color:'+ev.c+';color:'+ev.c:'')+'">'+ev.l+(isCur?' ●':'')+'</button>';}).join('');EK.forEach(k=>document.getElementById('p_'+k).style.display=k===cet?'block':'none');}
function swET(k){cet=k;EK.forEach(id=>{const btn=document.getElementById('et_'+id),isOn=id===k,isCur=id===cc?.curEvent;btn.classList.toggle('on',isOn);btn.style.background=isOn?EV[id].c:'';btn.style.color=isOn?'#fff':isCur?EV[id].c:'';btn.style.borderColor=isOn?EV[id].c:isCur?EV[id].c:'transparent';document.getElementById('p_'+id).style.display=isOn?'block':'none';});}
function hasEvData(r,ev){if(!r)return false;if(r.absent)return true;if(ev==='cardio')return r.cardio!==''&&r.cardio!=null;if(ev==='grip')return r.grip!==''&&r.grip!=null;if(ev==='jump')return r.jump!==''&&r.jump!=null;if(ev==='bmi')return !!(r.height&&r.weight);if(ev==='flex')return r.flex!==''&&r.flex!=null;return false;}
function gStat(r){if(!r)return 'none';if(r.absent)return 'absent';const ev=cc?.curEvent||'cardio';if(hasEvData(r,ev))return 'done';const any=[r.cardioGrade,r.gripGrade,r.jumpGrade,(r.height&&r.weight)?'1':'',r.flexGrade].some(v=>v!==''&&v!=null&&v!==undefined);return any?'part':'none';}
function getDone(grade,cls,ev){const event=ev||(cc?.curEvent)||'cardio';return records.filter(r=>r.grade===grade&&r.cls===cls&&(r.absent||hasEvData(r,event))).length;}
function updDone(){if(cc)document.getElementById('i_done_num').textContent=getDone(cc.grade,cc.cls,cc.curEvent);}
function findFirst(){if(!cc)return 0;for(let i=0;i<cc.count;i++){const r=getRec(i+1);if(!r||!hasEvData(r,cc.curEvent))return i;}return 0;}
function getRec(num){return records.find(r=>cc&&r.grade===cc.grade&&r.cls===cc.cls&&r.num===num);}
function startClass(){const grade=document.getElementById('s_grade').value,cls=document.getElementById('s_cls').value,evt=document.getElementById('s_event').value;if(!grade||!cls){showToast('학년과 반을 선택해주세요','#D85A30');return;}const key=grade+'_'+cls,ros=rosters[key]||[];let cl=classes.find(c=>c.key===key);if(!cl){cl={key,grade,cls,count:ros.length||30,date:document.getElementById('g_date').value,measurer:document.getElementById('g_msr').value,curEvent:evt};classes.push(cl);}else{cl.curEvent=evt||cl.curEvent;}cl.names=ros.map(r=>r.name||'');cl.genders=ros.map(r=>r.gender||'');cl.count=ros.length>0?ros.length:cl.count;saveAll();cc=cl;cet=cl.curEvent||'cardio';ci=findFirst();showScreen('roster');rRoster();rRecent();}
function rRoster(){const grid=document.getElementById('rc_grid');grid.innerHTML='';for(let i=1;i<=cc.count;i++){const r=getRec(i),st=gStat(r),isCur=i===ci+1,nm=r?.name||cc.names?.[i-1]||'';const cls2='rc'+(isCur?' cur':st==='done'?' done':st==='part'?' part':st==='absent'?' abs':'');const d=document.createElement('div');d.className=cls2;d.onclick=()=>{ci=i-1;loadInput();showScreen('input');};d.innerHTML='<div class="rcn">'+i+'</div><div class="rcm">'+nm+'</div>';grid.appendChild(d);}}
function startFF(){ci=0;loadInput();showScreen('input');}
function loadInput(){if(!cc)return;const num=ci+1,r=getRec(num),nm=r?.name||cc.names?.[ci]||'',gn=r?.gender||cc.genders?.[ci]||'';document.getElementById('i_circle').textContent=num;document.getElementById('i_nd').textContent=nm||num+'번';document.getElementById('i_sub').textContent=cc.grade+'학년 '+cc.cls+'반'+(EV[cc.curEvent]?' · '+EV[cc.curEvent].f:'');updDone();const pct=Math.round((num-1)/cc.count*100);document.getElementById('i_cn').textContent=num;document.getElementById('i_tn').textContent=cc.count;document.getElementById('i_pt').textContent=pct+'%';document.getElementById('i_pb').style.width=pct+'%';document.getElementById('i_tl').textContent=cc.count+'번';document.getElementById('i_nb').textContent=num<cc.count?'다음 →':'완료 ✓';const fids=['i_name','i_cardio','i_cardioG','i_gr1','i_gl1','i_gr2','i_gl2','i_grip','i_gripG','i_j1','i_j2','i_jump','i_jumpG','i_ht','i_wt','i_f1','i_f2','i_flex','i_flexG','i_memo'];fids.forEach(id=>document.getElementById(id).value='');document.getElementById('i_gen').value='';['grip_bb','jump_bb','flex_bb'].forEach(id=>document.getElementById(id).style.display='none');EK.forEach(k=>{const el=document.getElementById('gr_'+k);if(el){el.className='gr gn';el.style.background='';el.textContent='측정값을 입력하면 자동 계산됩니다';}});if(nm)document.getElementById('i_name').value=nm;if(gn)document.getElementById('i_gen').value=gn;if(r&&!r.absent){if(r.name)document.getElementById('i_name').value=r.name;if(r.gender)document.getElementById('i_gen').value=r.gender;const mp={cardio:'i_cardio',cardioGrade:'i_cardioG',gr_r1:'i_gr1',gr_l1:'i_gl1',gr_r2:'i_gr2',gr_l2:'i_gl2',grip:'i_grip',gripGrade:'i_gripG',jump1:'i_j1',jump2:'i_j2',jump:'i_jump',jumpGrade:'i_jumpG',height:'i_ht',weight:'i_wt',flex1:'i_f1',flex2:'i_f2',flex:'i_flex',flexGrade:'i_flexG',memo:'i_memo'};Object.entries(mp).forEach(([rk,fid])=>{if(r[rk]!==undefined&&r[rk]!=='')document.getElementById(fid).value=r[rk];});ag('cardio');agGrip();agJump();ag('bmi');agFlex();}else if(r?.absent){document.getElementById('i_memo').value='결석';}buildET();swET(cc.curEvent||'cardio');}
function collectInput(){const h=parseFloat(document.getElementById('i_ht').value),w=parseFloat(document.getElementById('i_wt').value);const bmi=h>0&&w>0?Math.round(w/((h/100)**2)*10)/10:'';const bmiG=bmi?bmi<18.5?3:bmi<23?1:bmi<25?2:bmi<30?4:5:'';return{id:cc.grade+'_'+cc.cls+'_'+(ci+1),grade:cc.grade,cls:cc.cls,num:ci+1,name:document.getElementById('i_name').value.trim(),gender:document.getElementById('i_gen').value,date:cc.date,measurer:cc.measurer,cardio:document.getElementById('i_cardio').value,cardioGrade:document.getElementById('i_cardioG').value,gr_r1:document.getElementById('i_gr1').value,gr_l1:document.getElementById('i_gl1').value,gr_r2:document.getElementById('i_gr2').value,gr_l2:document.getElementById('i_gl2').value,grip:document.getElementById('i_grip').value,gripGrade:document.getElementById('i_gripG').value,jump1:document.getElementById('i_j1').value,jump2:document.getElementById('i_j2').value,jump:document.getElementById('i_jump').value,jumpGrade:document.getElementById('i_jumpG').value,height:h||'',weight:w||'',bmi,bmiGrade:bmiG,flex1:document.getElementById('i_f1').value,flex2:document.getElementById('i_f2').value,flex:document.getElementById('i_flex').value,flexGrade:document.getElementById('i_flexG').value,memo:document.getElementById('i_memo').value,absent:false,confirmed:false};}
function saveNext(){const rec=collectInput();const idx=records.findIndex(r=>r.id===rec.id);if(idx>=0)records[idx]=rec;else records.push(rec);saveAll();updDone();showToast('✓ '+rec.num+'번 '+(EV[cc.curEvent]?.l||'')+' 저장');if(ci<cc.count-1){ci++;loadInput();}else{showDone();}}
function prevS(){if(ci>0){ci--;loadInput();}}
function nextS(){if(ci<cc.count-1){ci++;loadInput();}else showDone();}
function markAbs(){const id=cc.grade+'_'+cc.cls+'_'+(ci+1),nm=document.getElementById('i_name').value.trim()||cc.names?.[ci]||'';const rec={id,grade:cc.grade,cls:cc.cls,num:ci+1,name:nm,absent:true,date:cc.date,measurer:cc.measurer,memo:'결석',confirmed:false};const idx=records.findIndex(r=>r.id===id);if(idx>=0)records[idx]=rec;else records.push(rec);saveAll();updDone();showToast(ci+1+'번 결석 처리','#888780');if(ci<cc.count-1){ci++;loadInput();}else showDone();}
function skipS(){if(ci<cc.count-1){ci++;loadInput();}}
function showDone(){if(!cc)return;const ev=cc.curEvent||'cardio',evInfo=EV[ev];const clsRecs=records.filter(r=>r.grade===cc.grade&&r.cls===cc.cls).sort((a,b)=>a.num-b.num);const measured=clsRecs.filter(r=>!r.absent&&hasEvData(r,ev)).length,abs=clsRecs.filter(r=>r.absent).length,notDone=cc.count-measured-abs;document.getElementById('done_title').textContent=cc.grade+'학년 '+cc.cls+'반 — '+evInfo.f;document.getElementById('done_sub').textContent='총 '+cc.count+'명';document.getElementById('done_stats').innerHTML='<div class="dsc"><div class="dsl">측정완료</div><div class="dsv" style="color:#1D9E75">'+measured+'</div></div><div class="dsc"><div class="dsl">미완료</div><div class="dsv" style="color:#BA7517">'+notDone+'</div></div><div class="dsc"><div class="dsl">결석</div><div class="dsv" style="color:#D85A30">'+abs+'</div></div>';const cols=getDoneCols(ev);document.getElementById('done_ct').innerHTML='<i class="ti ti-list-check"></i> '+evInfo.f+' 결과';document.getElementById('done_thead').innerHTML='<tr><th style="width:28px">번</th><th style="width:52px">이름</th>'+cols.map(c=>'<th style="width:'+c.w+'">'+c.lbl+'</th>').join('')+'<th style="width:36px">상태</th></tr>';document.getElementById('done_body').innerHTML=Array.from({length:cc.count},(_,i)=>{const num=i+1,r=clsRecs.find(rr=>rr.num===num)||null,nm=r?.name||cc.names?.[i]||'';const isDone=r&&(r.absent||hasEvData(r,ev));const badge=r?.absent?'<span class="bdg ba">결석</span>':isDone?'<span class="bdg bd">완료</span>':'<span class="bdg bn">미완</span>';const cells=cols.map(c=>{const val=r?c.getVal(r):'',grade=r?c.getGrd(r):'';const gpHtml=grade?'<div class="gp gp'+grade+'" style="margin:0 auto">'+grade+'</div>':'<div class="gp gpX" style="margin:0 auto">—</div>';return '<td style="text-align:center">'+(val?'<div style="font-size:11px;font-weight:500;line-height:1.3">'+val+'</div>'+gpHtml:gpHtml)+'</td>';}).join('');return '<tr><td style="font-size:11px;font-weight:500">'+num+'</td><td style="font-size:11px">'+(nm||'—')+'</td>'+cells+'<td>'+badge+'</td></tr>';}).join('');showScreen('done');}
function getDoneCols(ev){if(ev==='cardio')return[{lbl:'횟수(회)',w:'52px',getVal:r=>r.cardio||'',getGrd:r=>r.cardioGrade||''}];if(ev==='grip')return[{lbl:'우1',w:'34px',getVal:r=>r.gr_r1||'',getGrd:r=>''},{lbl:'좌1',w:'34px',getVal:r=>r.gr_l1||'',getGrd:r=>''},{lbl:'우2',w:'34px',getVal:r=>r.gr_r2||'',getGrd:r=>''},{lbl:'좌2',w:'34px',getVal:r=>r.gr_l2||'',getGrd:r=>''},{lbl:'평균',w:'44px',getVal:r=>r.grip?r.grip+'kg':'',getGrd:r=>r.gripGrade||''}];if(ev==='jump')return[{lbl:'1회',w:'44px',getVal:r=>r.jump1||'',getGrd:r=>''},{lbl:'2회',w:'44px',getVal:r=>r.jump2||'',getGrd:r=>''},{lbl:'채택',w:'48px',getVal:r=>r.jump?r.jump+'cm':'',getGrd:r=>r.jumpGrade||''}];if(ev==='bmi')return[{lbl:'키',w:'44px',getVal:r=>r.height||'',getGrd:r=>''},{lbl:'몸무게',w:'48px',getVal:r=>r.weight?r.weight+'kg':'',getGrd:r=>''},{lbl:'BMI',w:'44px',getVal:r=>r.bmi||'',getGrd:r=>r.bmiGrade||''}];if(ev==='flex')return[{lbl:'1회',w:'44px',getVal:r=>r.flex1||'',getGrd:r=>''},{lbl:'2회',w:'44px',getVal:r=>r.flex2||'',getGrd:r=>''},{lbl:'채택',w:'48px',getVal:r=>r.flex?r.flex+'cm':'',getGrd:r=>r.flexGrade||''}];return[];}
function rRecent(){const wrap=document.getElementById('rw'),list=document.getElementById('rl');if(!classes.length){wrap.style.display='none';return;}wrap.style.display='block';list.innerHTML=classes.slice().reverse().map(c=>{const done=getDone(c.grade,c.cls,c.curEvent),ev=EV[c.curEvent],evL=ev?ev.f:'미설정',evC=ev?ev.c:'#888';return '<div class="cr" onclick="resumeC(\''+c.key+'\')"><div style="flex:1;min-width:0"><div style="font-size:14px;font-weight:500">'+c.grade+'학년 '+c.cls+'반 <span class="ep" style="background:'+evC+'22;color:'+evC+'">'+evL+'</span></div><div style="font-size:12px;color:#6b6b6b">'+done+'/'+c.count+'명 완료 · '+(c.date||'')+'</div></div><i class="ti ti-chevron-right" style="color:#9b9b9b"></i></div>';}).join('');}
function resumeC(key){cc=classes.find(c=>c.key===key);if(!cc)return;cet=cc.curEvent||'cardio';ci=findFirst();showScreen('roster');rRoster();}
function rSB(){const board=document.getElementById('status_board'),now=new Date();document.getElementById('sb_ts').textContent=now.getHours().toString().padStart(2,'0')+':'+now.getMinutes().toString().padStart(2,'0')+':'+now.getSeconds().toString().padStart(2,'0')+' 기준';if(!classes.length){board.innerHTML='<div style="text-align:center;padding:20px;color:#9b9b9b;font-size:13px">측정을 시작하면 여기에 표시됩니다</div>';return;}board.innerHTML=classes.map(c=>{const ev=EV[c.curEvent],done=getDone(c.grade,c.cls,c.curEvent),total=c.count,pct=total>0?Math.round(done/total*100):0;const pill=ev?'<span class="ep" style="background:'+ev.c+';color:#fff;font-size:12px;padding:5px 12px">'+ev.f+'</span>':'<span style="font-size:12px;color:#9b9b9b">종목 미설정</span>';const bc=pct>=100?'#1D9E75':'#378ADD';return '<div class="sbr"><div class="sbc">'+c.grade+'학년<br>'+c.cls+'반</div><div style="flex:1">'+pill+'<div style="height:6px;background:#ebebeb;border-radius:3px;margin-top:8px"><div style="height:100%;width:'+pct+'%;background:'+bc+';border-radius:3px;transition:width .4s"></div></div><div style="font-size:12px;color:#6b6b6b;margin-top:4px;font-weight:500;display:flex;align-items:center;gap:6px">'+(pct>=100?'✅':'')+'<span>'+done+'</span><span style="font-weight:400"> / '+total+'명 완료 ('+pct+'%)</span></div></div></div>';}).join('');}
function rDash(){const gStat2=r=>{if(!r)return 'none';if(r.absent)return 'absent';const f=[r.cardioGrade,r.gripGrade,r.jumpGrade,(r.height&&r.weight)?'1':'',r.flexGrade].filter(v=>v!==''&&v!=null);if(f.length>=5)return 'done';if(f.length>0)return 'partial';return 'none';};const t=records.length,done=records.filter(r=>gStat2(r)==='done').length,abs=records.filter(r=>r.absent).length,conf=records.filter(r=>r.confirmed).length;document.getElementById('d_stats').innerHTML='<div class="sc"><div class="sl">전체</div><div class="sv">'+t+'</div></div><div class="sc"><div class="sl">전종목완료</div><div class="sv" style="color:#1D9E75">'+done+'</div></div><div class="sc"><div class="sl">결석</div><div class="sv" style="color:#D85A30">'+abs+'</div></div><div class="sc"><div class="sl">교사확인</div><div class="sv" style="color:#185FA5">'+conf+'</div></div>';const fg=document.getElementById('df_g').value,fs=document.getElementById('df_s').value,q=document.getElementById('df_q').value.toLowerCase();let list=records.filter(r=>{if(fg&&r.grade+'학년'!==fg)return false;if(fs&&gStat2(r)!==fs)return false;if(q&&!(r.name||'').includes(q)&&!(r.cls||'').includes(q))return false;return true;}).sort((a,b)=>parseInt(a.grade)*10000+parseInt(a.cls)*100+parseInt(a.num)-(parseInt(b.grade)*10000+parseInt(b.cls)*100+parseInt(b.num)));const body=document.getElementById('d_body'),empty=document.getElementById('d_empty');if(!list.length){body.innerHTML='';empty.style.display='block';return;}empty.style.display='none';const gp=(g)=>g?'<div class="gp gp'+g+'" style="margin:0 auto">'+g+'</div>':'<div class="gp gpX" style="margin:0 auto">—</div>';body.innerHTML=list.map(r=>{const s=gStat2(r),ri=records.indexOf(r);const badge=s==='done'?'<span class="bdg bd">완료</span>':s==='absent'?'<span class="bdg ba">결석</span>':s==='partial'?'<span class="bdg bp2">일부</span>':'<span class="bdg bn">미완</span>';return '<tr><td style="font-size:10px">'+r.grade+'-'+r.cls+'-'+r.num+'</td><td style="font-weight:500;font-size:11px">'+(r.name||'—')+'</td><td style="text-align:center">'+gp(r.cardioGrade)+'</td><td style="text-align:center">'+gp(r.gripGrade)+'</td><td style="text-align:center">'+gp(r.jumpGrade)+'</td><td style="text-align:center">'+gp(r.bmiGrade)+'</td><td style="text-align:center">'+gp(r.flexGrade)+'</td><td>'+badge+'</td><td><div style="display:flex;gap:2px"><button class="ibt" onclick="openE('+ri+')"><i class="ti ti-edit"></i></button><button class="ibt" onclick="togC('+ri+')" style="'+(r.confirmed?'color:#1D9E75':'')+'"><i class="ti ti-check"></i></button><button class="ibt" onclick="delR('+ri+')" style="color:#D85A30"><i class="ti ti-trash"></i></button></div></td></tr>';}).join('');}
function togC(i){records[i].confirmed=!records[i].confirmed;saveAll();rDash();}
function delR(i){records.splice(i,1);saveAll();rDash();showToast('삭제 완료','#888780');}
function openE(i){ei=i;const r=records[i];document.getElementById('em_b').innerHTML='<div class="g2" style="gap:8px;margin-bottom:10px"><div class="fg"><label>이름</label><input type="text" id="m_nm" value="'+(r.name||'')+'"></div><div class="fg"><label>성별</label><select id="m_gn"><option value="">—</option><option value="M"'+(r.gender==='M'?' selected':'')+'>남</option><option value="F"'+(r.gender==='F'?' selected':'')+'>여</option></select></div><div class="fg"><label>심폐(회)</label><input type="number" id="m_c" value="'+(r.cardio||'')+'"></div><div class="fg"><label>심폐등급</label><input type="number" id="m_cg" value="'+(r.cardioGrade||'')+'" min="1" max="5"></div><div class="fg"><label>악력우1</label><input type="number" id="m_r1" value="'+(r.gr_r1||'')+'" step="0.1"></div><div class="fg"><label>악력좌1</label><input type="number" id="m_l1" value="'+(r.gr_l1||'')+'" step="0.1"></div><div class="fg"><label>악력우2</label><input type="number" id="m_r2" value="'+(r.gr_r2||'')+'" step="0.1"></div><div class="fg"><label>악력좌2</label><input type="number" id="m_l2" value="'+(r.gr_l2||'')+'" step="0.1"></div><div class="fg"><label>악력등급</label><input type="number" id="m_grg" value="'+(r.gripGrade||'')+'" min="1" max="5"></div><div class="fg"><label>제자리1회</label><input type="number" id="m_j1" value="'+(r.jump1||'')+'" step="0.1"></div><div class="fg"><label>제자리2회</label><input type="number" id="m_j2" value="'+(r.jump2||'')+'" step="0.1"></div><div class="fg"><label>순발등급</label><input type="number" id="m_jg" value="'+(r.jumpGrade||'')+'" min="1" max="5"></div><div class="fg"><label>키(cm)</label><input type="number" id="m_h" value="'+(r.height||'')+'" step="0.1"></div><div class="fg"><label>몸무게(kg)</label><input type="number" id="m_w" value="'+(r.weight||'')+'" step="0.1"></div><div class="fg"><label>유연1회</label><input type="number" id="m_f1" value="'+(r.flex1||'')+'" step="0.1"></div><div class="fg"><label>유연2회</label><input type="number" id="m_f2" value="'+(r.flex2||'')+'" step="0.1"></div><div class="fg"><label>유연등급</label><input type="number" id="m_fg" value="'+(r.flexGrade||'')+'" min="1" max="5"></div></div><div class="fg" style="margin-bottom:10px"><label>메모</label><input type="text" id="m_mm" value="'+(r.memo||'')+'"></div><button class="bp" onclick="saveE()">수정 저장</button>';document.getElementById('em').classList.add('open');}
function saveE(){const r=records[ei],h=parseFloat(document.getElementById('m_h').value),w=parseFloat(document.getElementById('m_w').value);r.name=document.getElementById('m_nm').value;r.gender=document.getElementById('m_gn').value;r.cardio=document.getElementById('m_c').value;r.cardioGrade=document.getElementById('m_cg').value;r.gr_r1=document.getElementById('m_r1').value;r.gr_l1=document.getElementById('m_l1').value;r.gr_r2=document.getElementById('m_r2').value;r.gr_l2=document.getElementById('m_l2').value;const bR=Math.max(parseFloat(r.gr_r1)||0,parseFloat(r.gr_r2)||0),bL=Math.max(parseFloat(r.gr_l1)||0,parseFloat(r.gr_l2)||0);r.grip=bR&&bL?Math.round((bR+bL)/2*10)/10:bR||bL;r.gripGrade=document.getElementById('m_grg').value;r.jump1=document.getElementById('m_j1').value;r.jump2=document.getElementById('m_j2').value;r.jump=Math.max(parseFloat(r.jump1)||0,parseFloat(r.jump2)||0)||'';r.jumpGrade=document.getElementById('m_jg').value;r.height=h||'';r.weight=w||'';if(h>0&&w>0){r.bmi=Math.round(w/((h/100)**2)*10)/10;r.bmiGrade=r.bmi<18.5?3:r.bmi<23?1:r.bmi<25?2:r.bmi<30?4:5;}r.flex1=document.getElementById('m_f1').value;r.flex2=document.getElementById('m_f2').value;r.flex=Math.max(parseFloat(r.flex1)||-999,parseFloat(r.flex2)||-999);if(r.flex===-999)r.flex='';r.flexGrade=document.getElementById('m_fg').value;r.memo=document.getElementById('m_mm').value;saveAll();closeM('em');rDash();showToast('수정 완료');}
function closeM(id){document.getElementById(id).classList.remove('open');}
function showToast(msg,color){const t=document.getElementById('toast');t.textContent=msg;t.style.background=color||'#1D9E75';t.style.display='block';setTimeout(()=>t.style.display='none',2200);}
loadAll();
document.getElementById('g_date').value=new Date().toISOString().split('T')[0];
rRecent();
showScreen('home');
</script>
</body>
</html>
