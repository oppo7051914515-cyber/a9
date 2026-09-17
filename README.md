<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>ระบบเช็กชื่อเข้าเรียนอัจฉริยะ (Smart Attendance System)</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Lucide Icons -->
    <script src="https://unpkg.com/lucide@latest"></script>
    <!-- QRCodeJS -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
    <!-- Export Libraries -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.31/jspdf.plugin.autotable.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

    <style>
        @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap');
        body { font-family: 'Sarabun', sans-serif; -webkit-tap-highlight-color: transparent; }
        .tab-content { transition: opacity 0.2s ease-in-out; }
    </style>
</head>
<body class="bg-slate-100 text-slate-800 min-h-screen flex flex-col md:flex-row pb-16 md:pb-0">

    <!-- Top Header (ครู - Mobile) -->
    <header id="mobileHeader" class="md:hidden bg-indigo-900 text-white p-4 sticky top-0 z-30 shadow-md flex justify-between items-center">
        <div class="flex items-center gap-2">
            <div class="p-1.5 bg-indigo-600 rounded-lg">
                <i data-lucide="scan-face" class="w-6 h-6"></i>
            </div>
            <div>
                <h1 class="font-bold text-sm leading-tight">ระบบเช็กชื่อเข้าเรียน</h1>
                <span id="mobileCurrentClassText" class="text-[11px] text-indigo-300">วิชา: CS101</span>
            </div>
        </div>
    </header>

    <!-- Sidebar Menu (ครู - Laptop/Desktop) -->
    <aside id="desktopSidebar" class="hidden md:flex w-72 bg-indigo-900 text-white p-5 flex-col justify-between shadow-xl flex-shrink-0">
        <div>
            <div class="flex items-center gap-3 mb-6">
                <div class="p-2.5 bg-indigo-600 rounded-xl shadow-lg">
                    <i data-lucide="scan-face" class="w-7 h-7"></i>
                </div>
                <div>
                    <h1 class="font-bold text-base leading-tight">ระบบเช็กชื่อใบหน้า</h1>
                    <span class="text-xs text-indigo-300">Smart Attendance v12.0</span>
                </div>
            </div>

            <!-- กล่องจัดการห้องเรียน / รายวิชา -->
            <div class="mb-6 bg-indigo-950/70 p-3.5 rounded-2xl border border-indigo-700/50 space-y-2.5">
                <label class="block text-xs text-indigo-300 font-semibold flex items-center gap-1">
                    <i data-lucide="book-open" class="w-3.5 h-3.5"></i> รายวิชา / ห้องเรียน
                </label>
                <select id="classroomSelect" onchange="changeClassroom()" class="w-full bg-indigo-800 text-white p-2.5 rounded-xl text-sm font-medium focus:outline-none focus:ring-2 focus:ring-indigo-400 border border-indigo-600">
                    <!-- JS Render Options -->
                </select>

                <button onclick="openAddClassroomModal()" class="w-full text-xs bg-indigo-600 hover:bg-indigo-500 text-white py-2 px-2 rounded-lg font-medium flex items-center justify-center gap-1 transition">
                    <i data-lucide="plus" class="w-3.5 h-3.5"></i> เพิ่มห้องเรียน / รายวิชาใหม่
                </button>
            </div>

            <!-- เมนูนำทางอาจารย์ -->
            <nav class="space-y-1.5">
                <button onclick="switchTab('session-control')" id="nav-session-control" class="nav-btn w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-medium bg-indigo-800 text-white shadow-md transition">
                    <i data-lucide="play-circle" class="w-4 h-4"></i> ตั้งค่าคาบเรียน & ควบคุม (ครู)
                </button>
                <button onclick="switchTab('scan-student')" id="nav-scan-student" class="nav-btn w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-medium text-indigo-200 hover:bg-indigo-800/60 transition">
                    <i data-lucide="smartphone" class="w-4 h-4"></i> หน้าเช็กชื่อนักเรียน
                </button>
                <button onclick="switchTab('students')" id="nav-students" class="nav-btn w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-medium text-indigo-200 hover:bg-indigo-800/60 transition">
                    <i data-lucide="user-plus" class="w-4 h-4"></i> จัดการรายชื่อนักเรียน
                </button>
                <button onclick="switchTab('edit-time')" id="nav-edit-time" class="nav-btn w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-medium text-indigo-200 hover:bg-indigo-800/60 transition">
                    <i data-lucide="clock" class="w-4 h-4"></i> ตารางบันทึกการเข้าเรียน
                </button>
            </nav>
        </div>
    </aside>

    <!-- Content Main Area -->
    <main class="flex-1 p-4 md:p-8 overflow-y-auto max-w-7xl mx-auto w-full">

        <!-- Page Header -->
        <div id="mainPageHeader" class="flex flex-col md:flex-row justify-between items-start md:items-center mb-6 pb-4 border-b border-slate-200 gap-3">
            <div>
                <h2 id="pageTitle" class="text-xl md:text-2xl font-bold text-slate-800">⏱️ ตั้งค่าคาบเรียน & ควบคุมการเช็กชื่อ</h2>
                <div class="flex items-center gap-2 mt-1 text-xs md:text-sm text-slate-500">
                    <span class="font-bold text-indigo-600 bg-indigo-50 px-2.5 py-0.5 rounded-md border border-indigo-100" id="currentClassText">CS101</span>
                </div>
            </div>
            <div class="flex items-center gap-2 bg-white shadow-sm border border-slate-200 text-indigo-700 px-3.5 py-1.5 rounded-xl text-xs font-semibold">
                <i data-lucide="calendar" class="w-4 h-4 text-indigo-600"></i>
                <span id="liveDateText"></span>
            </div>
        </div>

        <!-- 1. เมนู: กำหนดเวลาและคุมเปิด/ปิด (ครู) -->
        <section id="tab-session-control" class="tab-content space-y-6">
            <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
                <div class="lg:col-span-2 bg-white p-5 md:p-6 rounded-2xl shadow-sm border border-slate-200">
                    <h3 class="font-bold text-slate-800 text-base md:text-lg mb-4 flex items-center gap-2">
                        <i data-lucide="clock" class="w-5 h-5 text-indigo-600"></i> กำหนดเกณฑ์เวลาเข้าเรียนประจำคาบ
                    </h3>

                    <!-- ตั้งค่าช่วงเวลาตรง / สาย -->
                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-6">
                        <div class="bg-slate-50 p-4 rounded-2xl border border-slate-200">
                            <label class="block text-xs font-bold text-slate-700 mb-1.5">⏱️ สิ้นสุดเวลา "มาตรงเวลา"</label>
                            <input type="time" id="onTimeLimitInput" value="08:30" onchange="saveTimeRules()" class="w-full bg-white border border-slate-300 rounded-xl p-2.5 text-sm font-bold text-indigo-600 outline-none focus:ring-2 focus:ring-indigo-500">
                        </div>
                        <div class="bg-slate-50 p-4 rounded-2xl border border-slate-200">
                            <label class="block text-xs font-bold text-slate-700 mb-1.5">⚠️ สิ้นสุดเวลา "เข้าเรียนสาย"</label>
                            <input type="time" id="lateLimitInput" value="09:00" onchange="saveTimeRules()" class="w-full bg-white border border-slate-300 rounded-xl p-2.5 text-sm font-bold text-amber-600 outline-none focus:ring-2 focus:ring-amber-500">
                        </div>
                    </div>

                    <!-- นำเข้า Sync Code จากเครื่องนักเรียน -->
                    <div class="pt-4 border-t border-slate-100">
                        <label class="block text-xs font-bold text-indigo-900 mb-1.5">📥 กรอกรหัสหลักฐานจากนักเรียน (คัดลอกจากนักเรียนมาวางลงตาราง):</label>
                        <div class="flex gap-2">
                            <input type="text" id="pasteLogInput" placeholder="วางโค้ดหลักฐานการเช็กชื่อของนักเรียนที่นี่..." class="flex-1 border border-slate-300 rounded-xl px-3 py-2.5 text-xs outline-none focus:ring-2 focus:ring-indigo-500">
                            <button onclick="importStudentLogData()" class="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2.5 rounded-xl text-xs font-bold shadow flex items-center gap-1">
                                <i data-lucide="download" class="w-4 h-4"></i> ดึงลงตาราง
                            </button>
                        </div>
                    </div>
                </div>

                <div class="bg-white p-5 md:p-6 rounded-2xl shadow-sm border border-slate-200 flex flex-col justify-between">
                    <div>
                        <h3 class="font-bold text-slate-800 mb-4 pb-2 border-b flex items-center gap-2">
                            <i data-lucide="pie-chart" class="w-5 h-5 text-indigo-600"></i> สรุปผลคาบนี้
                        </h3>
                        <div class="space-y-3">
                            <div class="flex justify-between items-center p-3 bg-slate-50 rounded-xl">
                                <span class="text-xs font-semibold text-slate-600">ลงชื่อเข้าเรียนแล้ว</span>
                                <span id="sessionCheckedCount" class="font-bold text-indigo-600 text-base">0 คน</span>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </section>

        <!-- 2. หน้าเช็กชื่อสำหรับนักเรียนเท่านั้น (Student View) -->
        <section id="tab-scan-student" class="tab-content hidden space-y-6">
            <div id="shareControlBar" class="max-w-md mx-auto bg-indigo-900 text-white p-3.5 rounded-2xl flex items-center justify-between gap-2 shadow-md">
                <div class="flex items-center gap-2 font-semibold text-xs pl-1">
                    <i data-lucide="share-2" class="w-4 h-4 text-indigo-300"></i>
                    <span>แชร์ให้นักเรียน</span>
                </div>
                <div class="flex gap-1.5">
                    <button onclick="copyStudentLink()" class="bg-indigo-600 hover:bg-indigo-500 px-3 py-1.5 rounded-lg text-xs font-medium flex items-center gap-1 transition">
                        <i data-lucide="copy" class="w-3.5 h-3.5"></i> คัดลอกลิงก์
                    </button>
                    <button onclick="showQRCodeModal()" class="bg-amber-500 hover:bg-amber-400 text-slate-900 px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1 transition">
                        <i data-lucide="qr-code" class="w-3.5 h-3.5"></i> QR Code
                    </button>
                </div>
            </div>

            <!-- การ์ดเช็กชื่อนักเรียน -->
            <div class="max-w-md mx-auto bg-white p-6 rounded-3xl shadow-xl border border-slate-200">
                <div id="studentScanFormArea" class="space-y-4">
                    <div class="text-center mb-2">
                        <span id="studentScanSubjectBadge" class="bg-indigo-100 text-indigo-700 text-xs font-bold px-3 py-1 rounded-full inline-block mb-2">รายวิชา CS101</span>
                        <h3 class="text-xl font-bold text-slate-800">ลงชื่อเข้าเรียน</h3>
                        <p class="text-xs text-slate-500 mt-0.5">เลือกชื่อของคุณ และเปิดกล้องถ่ายภาพเพื่อยืนยันตัวตน</p>
                    </div>

                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1">1. เลือกชื่อ-รหัสนักเรียนของคุณ</label>
                        <select id="studentSelfSelect" class="w-full border border-slate-300 rounded-xl p-3 text-sm font-medium focus:ring-2 focus:ring-indigo-500 outline-none bg-slate-50">
                            <!-- JS Render Options -->
                        </select>
                    </div>

                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1">2. สแกนใบหน้าเข้าเรียน</label>
                        <div class="relative bg-slate-900 rounded-2xl aspect-square flex items-center justify-center overflow-hidden border-2 border-slate-200 shadow-inner">
                            <video id="videoStudent" class="w-full h-full object-cover hidden" autoplay playsinline></video>
                            <div id="studentCamPlaceholder" class="text-center p-6 text-slate-400">
                                <i data-lucide="camera" class="w-12 h-12 mx-auto mb-2 opacity-40"></i>
                                <p class="text-xs font-medium">กดปุ่ม "เปิดกล้อง" ด้านล่าง</p>
                            </div>
                        </div>
                    </div>

                    <div class="grid grid-cols-2 gap-2.5 pt-2">
                        <button onclick="startStudentCamera()" class="bg-indigo-600 hover:bg-indigo-700 active:scale-95 text-white py-3 rounded-xl text-xs md:text-sm font-bold flex items-center justify-center gap-1.5 shadow-md transition">
                            <i data-lucide="camera" class="w-4 h-4"></i> เปิดกล้อง
                        </button>
                        <button onclick="confirmStudentSelfScan()" class="bg-emerald-600 hover:bg-emerald-700 active:scale-95 text-white py-3 rounded-xl text-xs md:text-sm font-bold flex items-center justify-center gap-1.5 shadow-md transition">
                            <i data-lucide="check-circle" class="w-4 h-4"></i> บันทึกเช็กชื่อ
                        </button>
                    </div>
                </div>
            </div>
        </section>

        <!-- 3. เมนู: จัดการข้อมูลนักเรียน -->
        <section id="tab-students" class="tab-content hidden space-y-6">
            <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
                <!-- ฟอร์มเพิ่มนักเรียน -->
                <div class="bg-white p-5 rounded-2xl shadow-sm border border-slate-200">
                    <h3 class="font-bold text-slate-800 mb-4 flex items-center gap-2">
                        <i data-lucide="user-plus" class="w-5 h-5 text-indigo-600"></i> เพิ่มรายชื่อนักเรียนใหม่
                    </h3>
                    <form onsubmit="saveStudent(event)" class="space-y-3.5">
                        <div>
                            <label class="block text-xs font-semibold text-slate-600 mb-1">รหัสนักเรียน <span class="text-rose-500">*</span></label>
                            <input type="text" id="studentId" required class="w-full border border-slate-300 rounded-xl p-2.5 text-sm outline-none focus:ring-2 focus:ring-indigo-500" placeholder="เช่น 6601001">
                        </div>
                        <div>
                            <label class="block text-xs font-semibold text-slate-600 mb-1">ชื่อ-นามสกุล <span class="text-rose-500">*</span></label>
                            <input type="text" id="studentName" required class="w-full border border-slate-300 rounded-xl p-2.5 text-sm outline-none focus:ring-2 focus:ring-indigo-500" placeholder="เช่น นาย สมชาย ใจดี">
                        </div>
                        <div>
                            <label class="block text-xs font-semibold text-slate-600 mb-1">ชั้นเรียน / ห้อง</label>
                            <input type="text" id="studentClassGrade" class="w-full border border-slate-300 rounded-xl p-2.5 text-sm outline-none focus:ring-2 focus:ring-indigo-500" placeholder="เช่น ม.4/1">
                        </div>
                        <button type="submit" class="w-full bg-indigo-600 hover:bg-indigo-700 text-white py-2.5 rounded-xl text-sm font-semibold shadow transition">
                            บันทึกเพิ่มรายชื่อ
                        </button>
                    </form>
                </div>

                <!-- ตารางรายชื่อนักเรียน -->
                <div class="lg:col-span-2 bg-white p-5 rounded-2xl shadow-sm border border-slate-200">
                    <h3 class="font-bold text-slate-800 mb-4 flex items-center gap-2">
                        <i data-lucide="users" class="w-5 h-5 text-indigo-600"></i> รายชื่อนักเรียนในห้องนี้
                    </h3>
                    <div class="overflow-x-auto">
                        <table class="w-full text-left text-sm text-slate-600 border-collapse">
                            <thead class="bg-slate-100 text-slate-700 uppercase text-xs">
                                <tr>
                                    <th class="p-3 rounded-l-xl">รหัส</th>
                                    <th class="p-3">ชื่อ-นามสกุล</th>
                                    <th class="p-3">ชั้นเรียน</th>
                                    <th class="p-3 text-center rounded-r-xl">จัดการ</th>
                                </tr>
                            </thead>
                            <tbody id="studentListTable" class="divide-y divide-slate-100">
                                <!-- JS Render Student List -->
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </section>

        <!-- 4. เมนู: รายการเช็กชื่อ & Export -->
        <section id="tab-edit-time" class="tab-content hidden space-y-6">
            <div class="bg-white p-5 rounded-2xl shadow-sm border border-slate-200">
                <div class="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-4 gap-3">
                    <h3 class="font-bold text-slate-800 text-base flex items-center gap-2">
                        <i data-lucide="clock" class="w-5 h-5 text-indigo-600"></i> ตารางรายการเช็กชื่อเข้าเรียน
                    </h3>
                    <div class="flex gap-2">
                        <button onclick="exportPDF()" class="bg-rose-600 hover:bg-rose-700 text-white px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1 shadow">
                            <i data-lucide="file-text" class="w-3.5 h-3.5"></i> Export PDF
                        </button>
                        <button onclick="exportExcel()" class="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1 shadow">
                            <i data-lucide="sheet" class="w-3.5 h-3.5"></i> Export Excel
                        </button>
                    </div>
                </div>

                <div class="overflow-x-auto">
                    <table class="w-full text-left text-sm text-slate-600 border-collapse">
                        <thead class="bg-slate-100 text-slate-700 uppercase text-xs">
                            <tr>
                                <th class="p-3 rounded-l-xl">วันที่</th>
                                <th class="p-3">รหัส</th>
                                <th class="p-3">ชื่อ-นามสกุล</th>
                                <th class="p-3">เวลาที่ลงชื่อ</th>
                                <th class="p-3 rounded-r-xl">สถานะ</th>
                            </tr>
                        </thead>
                        <tbody id="attendanceLogsTable" class="divide-y divide-slate-100">
                            <!-- JS Render Logs -->
                        </tbody>
                    </table>
                </div>
            </div>
        </section>

    </main>

    <!-- Modal แสดงรหัสบันทึกสำหรับนักเรียน (นำส่งครู) -->
    <div id="receiptModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center hidden z-50 p-4">
        <div class="bg-white w-full max-w-sm rounded-3xl p-6 shadow-2xl border border-slate-200 text-center space-y-4">
            <div class="w-12 h-12 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center mx-auto">
                <i data-lucide="check-circle-2" class="w-7 h-7"></i>
            </div>
            <div>
                <h3 class="font-bold text-lg text-slate-800">เช็กชื่อสำเร็จ!</h3>
                <p id="receiptInfoText" class="text-xs text-slate-500 mt-1"></p>
            </div>

            <div class="bg-slate-50 p-3.5 rounded-2xl border border-slate-200 text-left">
                <label class="block text-[11px] font-bold text-slate-500 mb-1">📋 รหัสหลักฐานการเช็กชื่อ (ส่งให้คุณครู):</label>
                <input type="text" id="receiptCodeInput" readonly class="w-full bg-white border border-slate-300 rounded-xl p-2.5 text-xs font-mono font-bold text-indigo-700 text-center outline-none select-all">
            </div>

            <div class="space-y-2">
                <button onclick="copyReceiptCode()" class="w-full bg-indigo-600 hover:bg-indigo-700 text-white py-3 rounded-xl text-xs font-bold shadow flex items-center justify-center gap-1">
                    <i data-lucide="copy" class="w-4 h-4"></i> คัดลอกรหัสหลักฐาน
                </button>
                <button onclick="closeReceiptModal()" class="w-full bg-slate-100 hover:bg-slate-200 text-slate-600 py-2.5 rounded-xl text-xs font-semibold">
                    ปิดหน้าต่าง
                </button>
            </div>
        </div>
    </div>

    <!-- Bottom Nav มือถือ (ครู) -->
    <nav id="mobileBottomNav" class="md:hidden fixed bottom-0 left-0 right-0 bg-white border-t border-slate-200 flex justify-around items-center p-2 z-40 shadow-lg">
        <button onclick="switchTab('session-control')" class="flex flex-col items-center gap-1 text-[11px] font-semibold text-indigo-600">
            <i data-lucide="play-circle" class="w-5 h-5"></i> คุมเช็กชื่อ
        </button>
        <button onclick="switchTab('scan-student')" class="flex flex-col items-center gap-1 text-[11px] font-medium text-slate-400">
            <i data-lucide="smartphone" class="w-5 h-5"></i> หน้าสแกน
        </button>
        <button onclick="switchTab('students')" class="flex flex-col items-center gap-1 text-[11px] font-medium text-slate-400">
            <i data-lucide="user-plus" class="w-5 h-5"></i> นักเรียน
        </button>
        <button onclick="switchTab('edit-time')" class="flex flex-col items-center gap-1 text-[11px] font-medium text-slate-400">
            <i data-lucide="clock" class="w-5 h-5"></i> รายการ
        </button>
    </nav>

    <!-- Modal เพิ่มห้องเรียนใหม่ -->
    <div id="addClassroomModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center hidden z-50 p-4">
        <div class="bg-white w-full max-w-sm rounded-3xl p-6 shadow-2xl border border-slate-200">
            <h3 class="font-bold text-lg text-slate-800 mb-3">เพิ่มห้องเรียน / รายวิชาใหม่</h3>
            <div class="space-y-3">
                <div>
                    <label class="block text-xs font-semibold text-slate-600 mb-1">รหัสวิชา / รหัสห้อง <span class="text-rose-500">*</span></label>
                    <input type="text" id="newClassCode" placeholder="เช่น CS102 หรือ ม.4/1" class="w-full border border-slate-300 rounded-xl p-2.5 text-sm outline-none focus:ring-2 focus:ring-indigo-500">
                </div>
                <div>
                    <label class="block text-xs font-semibold text-slate-600 mb-1">ชื่อรายวิชา</label>
                    <input type="text" id="newClassSubject" placeholder="เช่น สังคมศึกษา" class="w-full border border-slate-300 rounded-xl p-2.5 text-sm outline-none focus:ring-2 focus:ring-indigo-500">
                </div>
            </div>
            <div class="flex gap-2 pt-4">
                <button onclick="saveNewClassroom()" class="flex-1 bg-indigo-600 text-white py-2.5 rounded-xl text-sm font-semibold">บันทึก</button>
                <button onclick="closeAddClassroomModal()" class="bg-slate-200 text-slate-700 px-4 py-2.5 rounded-xl text-sm font-semibold">ยกเลิก</button>
            </div>
        </div>
    </div>

    <!-- Modal QR Code -->
    <div id="qrCodeModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center hidden z-50 p-4">
        <div class="bg-white w-full max-w-sm rounded-3xl p-6 text-center shadow-2xl border border-slate-200">
            <h3 class="font-bold text-lg text-slate-800 mb-1">สแกน QR Code เพื่อเช็กชื่อ</h3>
            <p class="text-xs text-slate-500 mb-4">ให้นักเรียนใช้มือถือ/iPad สแกนเข้าสู่หน้าเช็กชื่อ</p>
            <div id="qrcode" class="flex justify-center p-4 bg-slate-50 rounded-2xl border border-slate-200 mx-auto mb-4"></div>
            <button onclick="closeQRCodeModal()" class="w-full bg-slate-800 hover:bg-slate-900 text-white py-3 rounded-xl text-sm font-semibold">
                ปิดหน้าต่าง
            </button>
        </div>
    </div>

    <script>
        let currentClassroom = 'CS101';

        let db = JSON.parse(localStorage.getItem('smart_attendance_db_v12')) || {
            timeRules: { onTimeLimit: '08:30', lateLimit: '09:00' },
            classrooms: {
                'CS101': 'CS101 - พัฒนาเว็บแอปพลิเคชัน',
                'SOC201': 'SOC201 - สังคมศึกษาและการสอน'
            },
            students: {
                'CS101': [
                    { id: '6601001', name: 'นาย กิตติศักดิ์ สมบูรณ์', grade: 'ม.4/1' },
                    { id: '6601002', name: 'นางสาว ปริศนา สุขใจ', grade: 'ม.4/1' },
                    { id: '6601003', name: 'นาย พัชรพล ต่อวาส', grade: 'ม.4/2' }
                ],
                'SOC201': [
                    { id: '6602001', name: 'นาย ณัฐวุฒิ มีสุข', grade: 'ปี 2' }
                ]
            },
            attendanceLogs: { 'CS101': [], 'SOC201': [] }
        };

        function saveData() {
            localStorage.setItem('smart_attendance_db_v12', JSON.stringify(db));
        }

        document.addEventListener("DOMContentLoaded", () => {
            lucide.createIcons();
            document.getElementById('liveDateText').innerText = new Date().toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' });

            if (db.timeRules) {
                document.getElementById('onTimeLimitInput').value = db.timeRules.onTimeLimit || '08:30';
                document.getElementById('lateLimitInput').value = db.timeRules.lateLimit || '09:00';
            }

            const urlParams = new URLSearchParams(window.location.search);
            const classParam = urlParams.get('class');

            if (classParam) {
                currentClassroom = classParam;
                ['mobileHeader', 'desktopSidebar', 'mobileBottomNav', 'mainPageHeader', 'shareControlBar'].forEach(id => {
                    const el = document.getElementById(id);
                    if (el) el.style.display = 'none';
                });
                document.body.classList.remove('pb-16', 'md:pb-0');
                switchTab('scan-student');
            }

            renderClassroomSelect();
            renderStudentSelectOptions();
            renderStudentList();
            renderAttendanceLogs();
        });

        function saveTimeRules() {
            const onTimeVal = document.getElementById('onTimeLimitInput').value;
            const lateVal = document.getElementById('lateLimitInput').value;
            db.timeRules = { onTimeLimit: onTimeVal, lateLimit: lateVal };
            saveData();
        }

        // --- ระบบเช็กชื่อ + สร้างรหัสหลักฐาน Sync ---
        function confirmStudentSelfScan() {
            const selectEl = document.getElementById('studentSelfSelect');
            if (!selectEl || !selectEl.value) {
                alert("กรุณาเลือกชื่อนักเรียนก่อนเช็กชื่อ");
                return;
            }

            const studentId = selectEl.value;
            const studentName = selectEl.selectedOptions[0].text;
            const now = new Date();

            const currentHour = String(now.getHours()).padStart(2, '0');
            const currentMinute = String(now.getMinutes()).padStart(2, '0');
            const currentTimeShort = `${currentHour}:${currentMinute}`;
            const timeFullStr = now.toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
            const dateStr = now.toISOString().split('T')[0];

            const onTimeLimit = db.timeRules?.onTimeLimit || '08:30';
            const lateLimit = db.timeRules?.lateLimit || '09:00';

            let status = "มาตรงเวลา";
            let statusColorClass = "bg-emerald-100 text-emerald-700";

            if (currentTimeShort > lateLimit) {
                status = "สายมาก / ขาด";
                statusColorClass = "bg-rose-100 text-rose-700";
            } else if (currentTimeShort > onTimeLimit) {
                status = "เข้าเรียนสาย";
                statusColorClass = "bg-amber-100 text-amber-700";
            }

            const newLog = {
                date: dateStr,
                id: studentId,
                name: studentName,
                time: timeFullStr,
                status: status,
                badgeClass: statusColorClass
            };

            if (!db.attendanceLogs[currentClassroom]) db.attendanceLogs[currentClassroom] = [];

            db.attendanceLogs[currentClassroom].unshift(newLog);
            saveData();
            renderAttendanceLogs();

            // สร้างรหัสหลักฐานส่งให้ครู
            const syncCode = btoa(unescape(encodeURIComponent(JSON.stringify(newLog))));
            document.getElementById('receiptCodeInput').value = syncCode;
            document.getElementById('receiptInfoText').innerText = `ชื่อ: ${studentName} | เวลา: ${timeFullStr} (${status})`;
            document.getElementById('receiptModal').classList.remove('hidden');
        }

        function copyReceiptCode() {
            const codeInput = document.getElementById('receiptCodeInput');
            codeInput.select();
            navigator.clipboard.writeText(codeInput.value).then(() => {
                alert("📋 คัดลอกรหัสหลักฐานเรียบร้อยแล้ว! นำรหัสนี้ส่งให้คุณครูได้เลยครับ");
            }).catch(() => {
                prompt("คัดลอกรหัสหลักฐานด้านล่างเพื่อส่งให้คุณครู:", codeInput.value);
            });
        }

        function closeReceiptModal() {
            document.getElementById('receiptModal').classList.add('hidden');
        }

        function importStudentLogData() {
            const code = document.getElementById('pasteLogInput').value.trim();
            if (!code) {
                alert("กรุณาวางรหัสหลักฐานจากนักเรียน");
                return;
            }

            try {
                const log = JSON.parse(decodeURIComponent(escape(atob(code))));
                if (!db.attendanceLogs[currentClassroom]) db.attendanceLogs[currentClassroom] = [];
                db.attendanceLogs[currentClassroom].unshift(log);
                saveData();
                renderAttendanceLogs();
                document.getElementById('pasteLogInput').value = '';
                alert(`✅ นำเข้าข้อมูลการเช็กชื่อของ ${log.name} เข้าสู่ตารางเรียบร้อยแล้ว!`);
            } catch (e) {
                alert("❌ รหัสหลักฐานไม่ถูกต้อง");
            }
        }

        function renderClassroomSelect() {
            const select = document.getElementById('classroomSelect');
            if (!select) return;
            select.innerHTML = Object.keys(db.classrooms).map(code => 
                `<option value="${code}" ${code === currentClassroom ? 'selected' : ''}>${db.classrooms[code]}</option>`
            ).join('');
        }

        function openAddClassroomModal() { document.getElementById('addClassroomModal').classList.remove('hidden'); }
        function closeAddClassroomModal() { document.getElementById('addClassroomModal').classList.add('hidden'); }

        function saveNewClassroom() {
            const code = document.getElementById('newClassCode').value.trim();
            const subject = document.getElementById('newClassSubject').value.trim();
            if (!code) return;
            db.classrooms[code] = `${code} - ${subject || 'วิชาทั่วไป'}`;
            if (!db.students[code]) db.students[code] = [];
            if (!db.attendanceLogs[code]) db.attendanceLogs[code] = [];
            currentClassroom = code;
            saveData();
            renderClassroomSelect();
            changeClassroom();
            closeAddClassroomModal();
        }

        function saveStudent(e) {
            e.preventDefault();
            const id = document.getElementById('studentId').value.trim();
            const name = document.getElementById('studentName').value.trim();
            const grade = document.getElementById('studentClassGrade').value.trim();

            if (!db.students[currentClassroom]) db.students[currentClassroom] = [];
            db.students[currentClassroom].push({ id, name, grade: grade || '-' });
            saveData();

            document.getElementById('studentId').value = '';
            document.getElementById('studentName').value = '';
            document.getElementById('studentClassGrade').value = '';

            renderStudentSelectOptions();
            renderStudentList();
            alert(`เพิ่มนักเรียน ${name} เรียบร้อยแล้ว`);
        }

        function deleteStudent(idx) {
            if (confirm("ต้องการลบรายชื่อนักเรียนคนนี้ออกหรือไม่?")) {
                db.students[currentClassroom].splice(idx, 1);
                saveData();
                renderStudentSelectOptions();
                renderStudentList();
            }
        }

        function renderStudentList() {
            const tbody = document.getElementById('studentListTable');
            if (!tbody) return;
            const students = db.students[currentClassroom] || [];
            tbody.innerHTML = students.map((s, idx) => `
                <tr class="hover:bg-slate-50">
                    <td class="p-3 font-medium">${s.id}</td>
                    <td class="p-3 font-semibold text-slate-800">${s.name}</td>
                    <td class="p-3 text-slate-500">${s.grade}</td>
                    <td class="p-3 text-center">
                        <button onclick="deleteStudent(${idx})" class="text-rose-600 hover:text-rose-800 p-1">
                            <i data-lucide="trash-2" class="w-4 h-4"></i>
                        </button>
                    </td>
                </tr>
            `).join('');
            lucide.createIcons();
        }

        function renderAttendanceLogs() {
            const tbody = document.getElementById('attendanceLogsTable');
            const countEl = document.getElementById('sessionCheckedCount');
            const logs = db.attendanceLogs[currentClassroom] || [];

            if (countEl) countEl.innerText = `${logs.length} คน`;
            if (!tbody) return;

            tbody.innerHTML = logs.map(l => {
                const badgeClass = l.badgeClass || (l.status === 'มาตรงเวลา' ? 'bg-emerald-100 text-emerald-700' : 'bg-amber-100 text-amber-700');
                return `
                    <tr class="hover:bg-slate-50">
                        <td class="p-3 font-medium">${l.date}</td>
                        <td class="p-3">${l.id}</td>
                        <td class="p-3 font-semibold text-slate-800">${l.name}</td>
                        <td class="p-3 font-mono text-indigo-600 font-bold">${l.time}</td>
                        <td class="p-3"><span class="px-2.5 py-0.5 rounded-full text-xs font-semibold ${badgeClass}">${l.status}</span></td>
                    </tr>
                `;
            }).join('');
        }

        function switchTab(tabId) {
            document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
            const target = document.getElementById(`tab-${tabId}`);
            if (target) target.classList.remove('hidden');
        }

        function changeClassroom() {
            currentClassroom = document.getElementById('classroomSelect').value;
            document.getElementById('currentClassText').innerText = currentClassroom;
            document.getElementById('studentScanSubjectBadge').innerText = `รายวิชา ${currentClassroom}`;
            renderStudentSelectOptions();
            renderStudentList();
            renderAttendanceLogs();
        }

        function renderStudentSelectOptions() {
            const select = document.getElementById('studentSelfSelect');
            if (!select) return;
            const students = db.students[currentClassroom] || [];
            select.innerHTML = students.map(s => `<option value="${s.id}">${s.name} (${s.id})</option>`).join('');
        }

        async function startStudentCamera() {
            const video = document.getElementById('videoStudent');
            const placeholder = document.getElementById('studentCamPlaceholder');
            try {
                const stream = await navigator.mediaDevices.getUserMedia({ video: true });
                video.srcObject = stream;
                video.classList.remove('hidden');
                placeholder.classList.add('hidden');
            } catch (err) {
                alert("ไม่สามารถเปิดกล้องได้: " + err.message);
            }
        }

        function getStudentPageURL() {
            let baseUrl = window.location.href.split('?')[0].split('#')[0];
            return `${baseUrl}?class=${encodeURIComponent(currentClassroom)}`;
        }

        function copyStudentLink() {
            const url = getStudentPageURL();
            navigator.clipboard.writeText(url).then(() => { alert("📋 คัดลอกลิงก์สแกนสำหรับนักเรียนเรียบร้อยแล้ว!\n" + url); });
        }

        function showQRCodeModal() {
            const url = getStudentPageURL();
            const qrContainer = document.getElementById("qrcode");
            qrContainer.innerHTML = "";
            new QRCode(qrContainer, { text: url, width: 180, height: 180 });
            document.getElementById('qrCodeModal').classList.remove('hidden');
        }

        function closeQRCodeModal() { document.getElementById('qrCodeModal').classList.add('hidden'); }

        function exportPDF() {
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF();
            doc.text(`Attendance Log - ${currentClassroom}`, 14, 15);
            const logs = db.attendanceLogs[currentClassroom] || [];
            doc.autoTable({
                head: [['Date', 'ID', 'Name', 'Time', 'Status']],
                body: logs.map(l => [l.date, l.id, l.name, l.time, l.status]),
                startY: 22
            });
            doc.save(`Attendance_${currentClassroom}.pdf`);
        }

        function exportExcel() {
            const logs = db.attendanceLogs[currentClassroom] || [];
            const worksheet = XLSX.utils.json_to_sheet(logs);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "Attendance");
            XLSX.writeFile(workbook, `Attendance_${currentClassroom}.xlsx`);
        }
    </script>
</body>
</html>
