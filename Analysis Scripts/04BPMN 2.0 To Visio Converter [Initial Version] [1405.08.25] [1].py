## ============================================================
## هدف و قابلیت‌های اصلی برنامه
## ============================================================
## این برنامه یک ابزار حرفه‌ای برای تبدیل دوطرفه فایل‌های BPMN و Microsoft Visio 2024 است.
##
## برنامه باید دارای یک منوی اصلی باشد و قابلیت‌های زیر را ارائه دهد:
##
## 1. تبدیل فایل BPMN با پسوند .bpmn به فایل واقعی Visio با پسوند .vsdx
## 2. تبدیل فایل Visio با پسوند .vsdx به فایل BPMN با پسوند .bpmn
## 3. بررسی ساختار فایل BPMN و شناسایی خطاها و ارجاعات نامعتبر
## 4. تهیه گزارش کامل از عملیات تبدیل، اطلاعات حفظ‌شده و محدودیت‌های احتمالی
##
## ============================================================
## الزامات محیط اجرا
## ============================================================
## سیستم‌عامل هدف: Windows 10 و Windows 11
## نسخه پایتون: Python 3.10 یا بالاتر
## نرم‌افزار موردنیاز: Microsoft Visio 2024 Desktop
## کتابخانه ارتباط با Visio: pywin32
##
## برنامه باید از طریق رابط COM با نسخه دسکتاپ Visio ارتباط برقرار کند.
## در صورت نصب نبودن Visio یا در دسترس نبودن COM، خطای مناسب نمایش داده شود.
##
## ============================================================
## بخش اول: تبدیل BPMN به Visio
## ============================================================
## فایل BPMN باید به‌صورت XML خوانده و ساختار آن با رعایت Namespaceها تحلیل شود.
##
## برنامه باید عناصر اصلی BPMN را شناسایی کند؛ از جمله:
## - Task و انواع مختلف فعالیت‌ها
## - Start Event و End Event
## - Intermediate Event
## - Exclusive Gateway و سایر Gatewayها
## - Call Activity
## - SubProcess
## - Sequence Flow
## - Message Flow
## - Association
## - Participant یا Pool
## - Lane و Lane Set
##
## خروجی باید یک فایل واقعی و قابل ویرایش Visio باشد، نه تصویر یا PDF.
##
## برای هر نوع عنصر BPMN باید تا حد امکان از شکل هندسی متناسب استفاده شود:
## - فعالیت‌ها: مستطیل با گوشه‌های گرد
## - رویدادها: دایره
## - دروازه‌ها: لوزی واقعی
## - زیرفرایندها: مستطیل با مشخصات بصری زیرفرایند
## - Pool و Lane: محدوده‌های مشخص و قابل تشخیص
##
## ارتباط بین عناصر باید با خطوط مناسب نمایش داده شود:
## - Sequence Flow: خط پیوسته همراه با فلش جهت
## - Message Flow: خط‌چین همراه با فلش مناسب
## - Association: خط نقطه‌چین یا سبک مناسب ارتباط
##
## در صورت وجود اطلاعات BPMN-DI، موقعیت، اندازه و مسیر خطوط باید از همان اطلاعات استخراج شود.
## مختصات عناصر و نقاط اتصال خطوط باید تا حد امکان مطابق نمودار اصلی حفظ شوند.
##
## تبدیل محور مختصات BPMN به مختصات Visio باید به‌درستی انجام شود.
## از آنجا که جهت محور عمودی در این دو محیط متفاوت است، مقدار Y باید اصلاح شود.
##
## Poolها و Laneها باید پیش از عناصر داخلی رسم شوند تا شکل‌های بزرگ، عناصر اصلی را نپوشانند.
## خطوط ارتباطی نیز باید پشت عناصر اصلی قرار بگیرند.
##
## ============================================================
## بخش دوم: تبدیل Visio به BPMN
## ============================================================
## برنامه باید فایل VSDX را به‌عنوان یک بسته XML فشرده بررسی کند.
##
## اطلاعات Shape Data هر شکل باید برای شناسایی نوع عنصر BPMN استفاده شود.
## برای هر شکل، در صورت وجود، اطلاعات زیر استخراج شود:
## - شناسه عنصر
## - نوع عنصر BPMN
## - عنوان عنصر
## - نوع فعالیت یا رویداد
## - شناسه فرایند مربوطه
## - شناسه عنصر مبدأ و مقصد برای خطوط ارتباطی
##
## خطوط ارتباطی نیز باید بررسی شوند تا نوع ارتباط و جهت آن مشخص شود.
##
## در صورت وجود اطلاعات کافی، عناصر شناسایی‌شده به ساختار معتبر BPMN تبدیل شوند.
## Sequence Flowها باید در داخل Process قرار بگیرند.
## Participantها و Message Flowها باید در ساختار Collaboration قرار بگیرند.
## Laneها و ارتباط عناصر با Lane مربوطه نیز تا حد امکان حفظ شوند.
##
## اگر فایل Visio فاقد اطلاعات معنایی BPMN باشد، برنامه فقط موارد قابل‌شناسایی را بازسازی کند.
## موارد نامشخص، ناقص یا غیرقابل تبدیل باید در گزارش نهایی ثبت شوند.
##
## ============================================================
## بخش سوم: آرشیو فایل اصلی BPMN
## ============================================================
## هنگام تبدیل BPMN به Visio، نسخه اصلی XML باید به‌صورت فشرده در فایل VSDX ذخیره شود.
##
## روش پیشنهادی:
## - فشرده‌سازی XML اصلی
## - محاسبه مقدار SHA256 برای بررسی صحت اطلاعات
## - ذخیره نسخه فشرده در بخش اختصاصی فایل VSDX
## - ذخیره اطلاعاتی مانند نام فایل اصلی، نسخه برنامه و زمان تبدیل
##
## برای افزایش اطمینان، در صورت امکان اطلاعات آرشیو در یک شکل مخفی یا بخش داده‌ای
## داخل سند Visio نیز نگهداری شود.
##
## هنگام تبدیل مجدد VSDX به BPMN، برنامه باید ابتدا وجود آرشیو اصلی را بررسی کند.
##
## کاربر باید بتواند یکی از دو روش زیر را انتخاب کند:
## 1. بازیابی دقیق فایل BPMN اصلی از آرشیو
## 2. بازسازی BPMN از وضعیت فعلی شکل‌ها و اطلاعات Visio
##
## توجه: بازیابی آرشیو، نسخه اولیه BPMN را برمی‌گرداند و تغییرات بعدی انجام‌شده
## در Visio را وارد فایل بازیابی‌شده نمی‌کند.
##
## ============================================================
## بخش چهارم: حفظ اطلاعات BPMN-DI
## ============================================================
## اطلاعات BPMN-DI شامل موقعیت و اندازه شکل‌ها و مسیر خطوط ارتباطی است.
##
## برنامه باید اطلاعات BPMNShape و BPMNEdge را استخراج کند.
## مختصات Bounds برای شکل‌ها و Waypoints برای خطوط باید حفظ شوند.
##
## اگر اطلاعات موقعیت برای بعضی عناصر وجود نداشت، برنامه باید برای آن عناصر
## یک چیدمان خودکار و منظم ایجاد کند.
##
## عناصر فاقد اطلاعات موقعیت نباید بدون بررسی روی عناصر موجود قرار بگیرند.
##
## زیرفرایندهایی که اطلاعات BPMN-DI مستقل ندارند، در صورت امکان باید در یک صفحه
## جداگانه از فایل Visio نمایش داده شوند.
##
## ============================================================
## بخش پنجم: اطلاعات اختصاصی شکل‌ها در Visio
## ============================================================
## برای اینکه فایل Visio در تبدیل مجدد قابل شناسایی باشد، هر شکل باید دارای Shape Data باشد.
##
## فیلدهای پیشنهادی:
## BPMN_ID
## BPMN_TYPE
## BPMN_NAME
## BPMN_KIND
## BPMN_SOURCE
## BPMN_TARGET
## BPMN_PROCESS
## BPMN_PAGE
##
## اطلاعات فوق باید با شناسه‌ها و ساختار واقعی BPMN هماهنگ باشند.
## برای خطوط ارتباطی نیز اطلاعات نوع ارتباط، مبدأ و مقصد ذخیره شود.
##
## ============================================================
## بخش ششم: اعتبارسنجی و کنترل خطا
## ============================================================
## برنامه باید قبل و بعد از تبدیل، کنترل‌های ساختاری مناسب انجام دهد.
##
## موارد زیر بررسی شوند:
## - شناسه‌های تکراری
## - ارجاع به عناصر ناموجود
## - مبدأ یا مقصد نامعتبر در جریان‌های ارتباطی
## - Processهای بدون ساختار معتبر
## - Participantهای دارای processRef نامعتبر
## - Laneهای دارای ارجاع نامعتبر
## - عناصر فاقد اطلاعات ضروری
## - ساختارهای پشتیبانی‌نشده یا مبهم
##
## خطاهای جدی باید از هشدارها تفکیک شوند.
## برنامه نباید در صورت وجود اطلاعات ناقص، موفقیت کامل و بدون نقص را اعلام کند.
##
## ============================================================
## بخش هفتم: گزارش عملیات
## ============================================================
## پس از هر عملیات، یک گزارش قابل خواندن تهیه شود.
##
## گزارش باید شامل موارد زیر باشد:
## - نام فایل ورودی
## - نام فایل خروجی
## - نوع عملیات انجام‌شده
## - تعداد عناصر شناسایی‌شده
## - تعداد ارتباطات شناسایی‌شده
## - تعداد Poolها و Laneها
## - تعداد عناصر تبدیل‌شده
## - تعداد عناصر تبدیل‌نشده
## - خطاها و هشدارها
## - وضعیت آرشیو فایل اصلی
## - وضعیت اعتبارسنجی خروجی
##
## گزارش می‌تواند با فرمت TXT یا JSON ذخیره شود.
##
## ============================================================
## بخش هشتم: الزامات فنی و پایداری
## ============================================================
## تمام عملیات مربوط به Visio باید با مدیریت صحیح COM انجام شوند.
## در صورت بروز خطا، منابع COM باید تا حد امکان آزاد شوند.
##
## دستورات حساس Visio باید در بلوک‌های مدیریت خطا قرار بگیرند.
## برنامه نباید به نام‌های ثابت یا شناسه‌های غیرقابل‌اعتماد شکل‌ها وابسته باشد.
##
## برای جلوگیری از بازنویسی فایل‌های ورودی یا خروجی‌های قبلی، نام خروجی باید
## به‌صورت ایمن و با پسوند مشخص ساخته شود.
##
## نمونه نام‌گذاری:
## filename_converted.vsdx
## filename_reconstructed.bpmn
##
## تمام توابع باید ساختارمند، قابل نگهداری و دارای توضیحات مناسب باشند.
## بخش‌های اصلی برنامه باید از یکدیگر جدا باشند:
## - خواندن و تحلیل BPMN
## - استخراج اطلاعات BPMN-DI
## - ایجاد نمودار Visio
## - استخراج اطلاعات VSDX
## - بازسازی BPMN
## - اعتبارسنجی
## - تهیه گزارش
## - مدیریت منوی اصلی
##
## ============================================================
## محدودیت مهم
## ============================================================
## تبدیل کاملاً بدون اتلاف اطلاعات بین BPMN و Visio برای تمام فایل‌ها قابل تضمین نیست.
##
## بازیابی دقیق فایل اصلی فقط در صورتی امکان‌پذیر است که آرشیو اصلی در فایل VSDX
## وجود داشته و سالم باقی مانده باشد.
##
## بازسازی معنای BPMN از یک فایل Visio دلخواه نیز فقط زمانی امکان‌پذیر است
## که اطلاعات کافی در Shape Data و ساختار ارتباطات آن وجود داشته باشد.
##
## بنابراین برنامه باید اطلاعات قابل تبدیل را حفظ کند و تمام محدودیت‌ها و
## موارد پشتیبانی‌نشده را به‌صورت شفاف در گزارش اعلام کند.
##
## ============================================================
## نتیجه مورد انتظار
## ============================================================
## خروجی نهایی باید یک ابزار قابل استفاده، پایدار، قابل توسعه و مناسب ارائه حرفه‌ای باشد.
## نمودارهای Visio باید قابل ویرایش، خوانا و تا حد امکان منطبق با استاندارد BPMN باشند.
## فایل‌های BPMN خروجی باید ساختار XML معتبر و قابل استفاده در ابزارهای BPMN داشته باشند.
## ============================================================
##
## ============================================================
## PROJECT OBJECTIVE AND CORE FEATURES
## ============================================================
## This application is a professional bidirectional conversion tool for BPMN
## files and Microsoft Visio 2024 documents.
##
## The application must provide a main menu with the following options:
##
## 1. Convert BPMN files (.bpmn) into native, editable Visio files (.vsdx).
## 2. Convert Visio files (.vsdx) into BPMN files (.bpmn).
## 3. Validate BPMN document structure and detect errors and invalid references.
## 4. Generate detailed conversion reports, including preserved data and limitations.
##
## ============================================================
## RUNTIME ENVIRONMENT REQUIREMENTS
## ============================================================
## Target operating systems: Windows 10 and Windows 11.
## Required Python version: Python 3.10 or later.
## Required application: Microsoft Visio 2024 Desktop.
## Required Visio integration library: pywin32.
##
## The application must communicate with the installed Visio desktop application
## through the Windows COM automation interface.
##
## If Visio is not installed or its COM interface is unavailable, the application
## must display a clear and appropriate error message.
##
## ============================================================
## PART 1: BPMN TO VISIO CONVERSION
## ============================================================
## BPMN files must be read as XML documents and parsed using namespace-aware methods.
##
## The application must identify the main BPMN elements, including:
## - Tasks and different activity types.
## - Start Events and End Events.
## - Intermediate Events.
## - Exclusive Gateways and other Gateway types.
## - Call Activities.
## - SubProcesses.
## - Sequence Flows.
## - Message Flows.
## - Associations.
## - Participants and Pools.
## - Lanes and Lane Sets.
##
## The output must be a genuine, editable Visio document, not an image or PDF.
##
## Each BPMN element should be represented by an appropriate geometric shape:
## - Activities: rounded rectangles.
## - Events: circles.
## - Gateways: actual diamond shapes.
## - SubProcesses: rectangles with recognizable subprocess notation.
## - Pools and Lanes: clearly defined and identifiable regions.
##
## Connections between elements must use appropriate line styles:
## - Sequence Flow: solid line with a directional arrow.
## - Message Flow: dashed line with an appropriate arrow.
## - Association: dotted or otherwise appropriate association line.
##
## If BPMN-DI information is available, the original element positions,
## dimensions, and connection paths must be extracted and preserved.
##
## Shape coordinates and connector waypoints should remain as close as possible
## to the original BPMN diagram.
##
## The BPMN coordinate system must be converted correctly into Visio coordinates.
## Since the vertical axis direction differs between the two systems,
## the Y coordinate must be transformed appropriately.
##
## Pools and Lanes must be drawn before their internal elements so that
## large background shapes do not cover the actual BPMN elements.
##
## Connectors should also be placed behind the main diagram shapes.
##
## ============================================================
## PART 2: VISIO TO BPMN CONVERSION
## ============================================================
## The application must inspect VSDX documents as compressed XML-based packages.
##
## Shape Data must be used to identify the BPMN meaning of individual Visio shapes.
##
## Whenever available, the following information must be extracted:
## - Element ID.
## - BPMN element type.
## - Element name.
## - Activity or event subtype.
## - Related process ID.
## - Source and target element IDs for connectors.
##
## Connectors must also be inspected to determine their connection type and direction.
##
## When sufficient semantic information is available, the identified elements
## must be reconstructed as valid BPMN structures.
##
## Sequence Flows must be placed inside their corresponding Process elements.
## Participants and Message Flows must be represented within a Collaboration.
## Lane structures and element-to-lane references should be preserved whenever possible.
##
## If a Visio document does not contain sufficient BPMN metadata, the application
## must reconstruct only the elements that can be reliably identified.
##
## Ambiguous, incomplete, or unsupported elements must be documented in the report.
##
## ============================================================
## PART 3: ORIGINAL BPMN DOCUMENT ARCHIVING
## ============================================================
## During BPMN-to-Visio conversion, the original BPMN XML document must be
## stored in compressed form inside the resulting VSDX package.
##
## Recommended implementation:
## - Compress the original BPMN XML.
## - Calculate a SHA256 checksum to verify data integrity.
## - Store the compressed archive in a dedicated location inside the VSDX package.
## - Store metadata such as the original filename, application version,
##   and conversion timestamp.
##
## For additional reliability, archive information may also be stored in
## a hidden Visio shape or another document-level data section, if practical.
##
## When converting the VSDX file back to BPMN, the application must first
## check whether the original BPMN archive is available.
##
## The user must be able to choose between two conversion modes:
##
## 1. Restore the original BPMN XML from the embedded archive.
## 2. Reconstruct BPMN from the current Visio shapes and their metadata.
##
## Important: Restoring the archive returns the original BPMN document.
## It does not automatically include subsequent edits made in Visio.
##
## ============================================================
## PART 4: BPMN-DI PRESERVATION
## ============================================================
## BPMN-DI contains the graphical layout information of BPMN diagrams,
## including shape positions, dimensions, and connector paths.
##
## The application must extract both BPMNShape and BPMNEdge information.
##
## Shape Bounds and connector Waypoints must be preserved whenever available.
##
## If graphical layout information is missing for certain elements,
## the application must generate a clean and consistent automatic layout.
##
## Elements without layout information must not be placed over existing shapes
## without checking for possible overlaps.
##
## SubProcesses without independent BPMN-DI information should be displayed
## on a separate Visio page whenever practical.
##
## ============================================================
## PART 5: VISIO SHAPE METADATA
## ============================================================
## To ensure that Visio shapes can be identified during reverse conversion,
## each relevant shape must contain appropriate Shape Data.
##
## Recommended metadata fields:
## BPMN_ID
## BPMN_TYPE
## BPMN_NAME
## BPMN_KIND
## BPMN_SOURCE
## BPMN_TARGET
## BPMN_PROCESS
## BPMN_PAGE
##
## These fields must correspond to the actual BPMN element IDs and structure.
##
## Connectors must also store their connection type, source ID, and target ID.
##
## ============================================================
## PART 6: VALIDATION AND ERROR HANDLING
## ============================================================
## The application must perform appropriate structural validation
## before and after conversion.
##
## The following conditions must be checked:
## - Duplicate element IDs.
## - References to nonexistent elements.
## - Invalid source or target references in flows.
## - Processes with invalid structures.
## - Participants with invalid processRef values.
## - Lanes with invalid element references.
## - Elements missing required information.
## - Unsupported or ambiguous BPMN constructs.
##
## Critical errors must be distinguished from non-critical warnings.
##
## The application must never report a completely successful conversion
## when important information is missing or has not been preserved.
##
## ============================================================
## PART 7: CONVERSION REPORTING
## ============================================================
## A readable report must be generated after every conversion operation.
##
## The report must include:
## - Input filename.
## - Output filename.
## - Conversion operation type.
## - Number of identified elements.
## - Number of identified connections.
## - Number of Pools and Lanes.
## - Number of successfully converted elements.
## - Number of elements that could not be converted.
## - Errors and warnings.
## - Original document archive status.
## - Output validation status.
##
## Reports may be saved in TXT or JSON format.
##
## ============================================================
## PART 8: TECHNICAL RELIABILITY AND MAINTAINABILITY
## ============================================================
## All Visio automation operations must use proper COM resource management.
##
## COM resources should be released as safely as possible, including when errors occur.
##
## Sensitive Visio API operations must be protected by appropriate exception handling.
##
## The application must not depend on fixed shape names or unreliable shape identifiers.
##
## Output filenames must be generated safely to avoid overwriting source files
## or previously generated documents.
##
## Recommended output naming conventions:
## filename_converted.vsdx
## filename_reconstructed.bpmn
##
## All functions must be structured, maintainable, and appropriately documented.
##
## The main application components must be separated into independent modules
## or clearly defined functional sections:
##
## - BPMN parsing and analysis.
## - BPMN-DI extraction.
## - Visio diagram generation.
## - VSDX inspection and metadata extraction.
## - BPMN reconstruction.
## - Structural validation.
## - Conversion reporting.
## - Main menu and application workflow.
##
## ============================================================
## IMPORTANT TECHNICAL LIMITATIONS
## ============================================================
## Fully lossless conversion between BPMN and Visio cannot be guaranteed
## for every possible document.
##
## Exact restoration of the original BPMN file is possible only if the original
## archive exists inside the VSDX document and remains intact.
##
## Reconstructing BPMN semantics from an arbitrary Visio document is possible
## only when sufficient Shape Data and connection information are available.
##
## Therefore, the application must preserve all reliably convertible information
## and clearly report unsupported elements, ambiguities, and conversion limitations.
##
## ============================================================
## EXPECTED FINAL RESULT
## ============================================================
## The final product must be a reliable, maintainable, extensible,
## and professionally presentable application.
##
## Generated Visio diagrams must be editable, readable, and as BPMN-compliant
## as reasonably possible.
##
## Generated BPMN files must contain valid XML structures and remain usable
## in compatible BPMN modeling applications.
## ============================================================

"""
BPMN 2.0 TO MICROSOFT VISIO CONVERTER
Version: 1.0

Purpose:
    Convert BPMN 2.0 XML files (.bpmn / .xml) into editable Visio files (.vsdx).

Features:
    - Reads BPMN semantic elements and BPMN-DI graphical information.
    - Preserves original positions and dimensions whenever possible.
    - Automatically lays out elements when graphical information is missing.
    - Detects overlapping nodes and applies a basic layout correction.
    - Draws BPMN events, activities, gateways, pools, lanes, data objects,
      annotations, groups, and sequence/message/association flows.
    - Improves text readability using automatic font sizing and wrapping.
    - Supports Persian and English text.
    - Adds a conversion report.
    - Uses Microsoft Visio Desktop through COM automation.

Requirements:
    Python 3.10+
    Microsoft Visio Desktop
    pip install pywin32 lxml

Important:
    This is a graphical converter, not a complete BPMN execution engine.
    Unsupported or ambiguous BPMN constructs are reported.
"""


from __future__ import annotations

import os
import re
import sys
import math
import unicodedata
import traceback
import xml.etree.ElementTree as ET

from pathlib import Path
from collections import defaultdict

import win32com.client
import arabic_reshaper
from bidi.algorithm import get_display


APP_NAME = "BPMN 2.0 to Visio Converter"
VERSION = "1.1"

# BPMN coordinates are generally expressed in diagram units.
BPMN_TO_INCH = 1.0 / 96.0

# General page and layout settings.
PAGE_MARGIN = 0.45
NODE_WIDTH = 1.55
NODE_HEIGHT = 0.75
HORIZONTAL_GAP = 0.65
VERTICAL_GAP = 0.65

FONT_NAME = "Segoe UI"
MIN_FONT_SIZE = 7
MAX_FONT_SIZE = 11

# BPMN element groups.
FLOW_NODE_TYPES = {
    "startEvent", "endEvent",
    "intermediateThrowEvent", "intermediateCatchEvent",
    "boundaryEvent",
    "task", "userTask", "manualTask", "serviceTask",
    "scriptTask", "businessRuleTask", "sendTask", "receiveTask",
    "subProcess", "transaction", "callActivity",
    "exclusiveGateway", "parallelGateway", "inclusiveGateway",
    "eventBasedGateway", "complexGateway",
    "choreographyTask", "callChoreography", "subChoreography",
}

CONTAINER_TYPES = {
    "participant", "lane", "laneSet", "process",
    "subProcess", "transaction",
}

DATA_TYPES = {
    "dataObject", "dataObjectReference",
    "dataStore", "dataStoreReference",
    "inputDataItem", "outputDataItem",
}

FLOW_TYPES = {
    "sequenceFlow", "messageFlow", "association",
    "dataInputAssociation", "dataOutputAssociation",
    "conversationLink",
}

ANNOTATION_TYPES = {
    "textAnnotation", "group",
}

EVENT_TYPES = {
    "startEvent", "endEvent",
    "intermediateThrowEvent", "intermediateCatchEvent",
    "boundaryEvent",
}

GATEWAY_TYPES = {
    "exclusiveGateway", "parallelGateway",
    "inclusiveGateway", "eventBasedGateway",
    "complexGateway",
}

TASK_TYPES = {
    "task", "userTask", "manualTask", "serviceTask",
    "scriptTask", "businessRuleTask", "sendTask",
    "receiveTask", "callActivity", "subProcess",
    "transaction", "choreographyTask",
    "callChoreography", "subChoreography",
}

REPORT = {
    "warnings": [],
    "errors": [],
    "converted_elements": 0,
    "converted_flows": 0,
    "pages": 0,
}


# ============================================================
# PERSIAN / ARABIC TEXT PROCESSING
# ============================================================

# Characters commonly used in Arabic text but normally expected
# in Persian text. ZWNJ is deliberately preserved.
PERSIAN_CHARACTER_MAP = str.maketrans({
    "\u064A": "\u06CC",  # Arabic Yeh -> Persian Yeh
    "\u0649": "\u06CC",  # Alef Maksura -> Persian Yeh
    "\u0643": "\u06A9",  # Arabic Kaf -> Persian Keheh
})


def contains_arabic_script(text):
    """Return True if text contains Arabic or Persian script characters."""
    if not text:
        return False

    for char in text:
        code = ord(char)

        if (
            0x0600 <= code <= 0x06FF
            or 0x0750 <= code <= 0x077F
            or 0x08A0 <= code <= 0x08FF
            or 0xFB50 <= code <= 0xFDFF
            or 0xFE70 <= code <= 0xFEFF
        ):
            return True

    return False


def normalize_persian_text(text):
    """
    Normalize Unicode text without removing Persian characters
    or the Zero Width Non-Joiner (ZWNJ).
    """
    text = unicodedata.normalize("NFC", text or "")
    text = text.translate(PERSIAN_CHARACTER_MAP)

    return text


def prepare_visio_display_text(text):
    """
    Prepare text for visual display in Microsoft Visio.

    Processing steps:
    1. Unicode NFC normalization.
    2. Persian character normalization.
    3. Arabic/Persian glyph shaping.
    4. Bidirectional text reordering with RTL as base direction.

    Lines are processed independently so that separate labels,
    markers and multiline text are less likely to interfere
    with each other's direction.

    IMPORTANT:
    The returned string is intended for visual display in Visio.
    Do not write it back into the original BPMN XML.
    """
    if text is None:
        return ""

    text = normalize_persian_text(str(text))

    if not contains_arabic_script(text):
        return text

    processed_lines = []

    for line in text.splitlines():

        if not contains_arabic_script(line):
            processed_lines.append(line)
            continue

        try:
            reshaped_line = arabic_reshaper.reshape(line)

            visual_line = get_display(
                reshaped_line,
                base_dir="R",
            )

            processed_lines.append(visual_line)

        except Exception as exc:
            REPORT["warnings"].append(
                f"Persian text processing warning: {exc}"
            )

            # Preserve the original line if shaping fails.
            processed_lines.append(line)

    return "\n".join(processed_lines)


# ============================================================
# GENERAL XML UTILITIES
# ============================================================

def local_name(tag):
    """Return an XML tag name without its namespace."""
    if not isinstance(tag, str):
        return ""

    return tag.split("}")[-1]


def get_attr(element, name, default=""):
    return element.attrib.get(name, default)


def get_children(element, name):
    return [
        child for child in list(element)
        if local_name(child.tag) == name
    ]


def find_children_recursive(element, accepted_names):
    """Find elements recursively by their local XML names."""
    result = []

    for child in element.iter():
        if local_name(child.tag) in accepted_names:
            result.append(child)

    return result


def get_document_elements(root):
    """Create a global index of XML elements by their BPMN IDs."""
    result = {}

    for element in root.iter():
        element_id = get_attr(element, "id")

        if element_id:
            result[element_id] = element

    return result


def get_diagram_information(root):
    """
    Read BPMNDiagram, BPMNPlane, BPMNShape and BPMNEdge information.
    """
    shapes = {}
    edges = {}
    diagrams = []

    for element in root.iter():

        if local_name(element.tag) != "BPMNDiagram":
            continue

        plane = None

        for child in element:
            if local_name(child.tag) == "BPMNPlane":
                plane = child
                break

        if plane is None:
            continue

        diagram_shapes = {}
        diagram_edges = {}

        for item in plane:

            item_type = local_name(item.tag)
            bpmn_element = get_attr(item, "bpmnElement")

            if not bpmn_element:
                continue

            if item_type == "BPMNShape":

                bounds = None

                for child in item:
                    if local_name(child.tag) == "Bounds":
                        bounds = child
                        break

                if bounds is not None:
                    diagram_shapes[bpmn_element] = {
                        "x": float(get_attr(bounds, "x", 0)),
                        "y": float(get_attr(bounds, "y", 0)),
                        "width": float(get_attr(bounds, "width", 100)),
                        "height": float(get_attr(bounds, "height", 60)),
                    }

            elif item_type == "BPMNEdge":

                points = []

                for child in item:
                    if local_name(child.tag) == "waypoint":
                        points.append({
                            "x": float(get_attr(child, "x", 0)),
                            "y": float(get_attr(child, "y", 0)),
                        })

                diagram_edges[bpmn_element] = points

        diagrams.append({
            "shapes": diagram_shapes,
            "edges": diagram_edges,
            "plane_element": get_attr(plane, "bpmnElement"),
        })

        shapes.update(diagram_shapes)
        edges.update(diagram_edges)

    return diagrams, shapes, edges


# ============================================================
# BPMN ELEMENT CLASSIFICATION
# ============================================================

def is_flow_node(element):
    return local_name(element.tag) in FLOW_NODE_TYPES


def is_data_element(element):
    return local_name(element.tag) in DATA_TYPES


def is_annotation(element):
    return local_name(element.tag) in ANNOTATION_TYPES


def is_flow_element(element):
    return local_name(element.tag) in FLOW_TYPES


def is_container(element):
    return local_name(element.tag) in {
        "participant",
        "lane",
    }


def get_element_name(element):
    name = get_attr(element, "name").strip()

    if name:
        return name

    return ""


def get_event_definition(element):
    """
    Identify event definitions such as:
    Message, Timer, Error, Signal, Conditional, Escalation, etc.
    """
    definitions = []

    for child in element:
        name = local_name(child.tag)

        if name.endswith("EventDefinition"):
            definitions.append(
                name.replace("EventDefinition", "")
            )

    return definitions


def get_gateway_marker(element):
    gateway_type = local_name(element.tag)

    markers = {
        "exclusiveGateway": "X",
        "parallelGateway": "+",
        "inclusiveGateway": "O",
        "eventBasedGateway": "E",
        "complexGateway": "*",
    }

    return markers.get(gateway_type, "")


# ============================================================
# LAYOUT UTILITIES
# ============================================================

def rectangles_overlap(a, b, padding=0.05):
    return not (
        a["x"] + a["width"] + padding <= b["x"]
        or b["x"] + b["width"] + padding <= a["x"]
        or a["y"] + a["height"] + padding <= b["y"]
        or b["y"] + b["height"] + padding <= a["y"]
    )


def auto_layout(elements):
    """
    Basic automatic layout for elements without BPMN-DI information.
    Elements are arranged in rows.
    """
    positions = {}

    if not elements:
        return positions

    columns = max(3, math.ceil(math.sqrt(len(elements))))

    for index, element in enumerate(elements):

        row = index // columns
        column = index % columns

        element_id = get_attr(element, "id")

        positions[element_id] = {
            "x": PAGE_MARGIN + column * (
                NODE_WIDTH + HORIZONTAL_GAP
            ),
            "y": PAGE_MARGIN + row * (
                NODE_HEIGHT + VERTICAL_GAP
            ),
            "width": NODE_WIDTH,
            "height": NODE_HEIGHT,
        }

    return positions


def correct_overlaps(positions, elements):
    """
    Move automatically positioned nodes if their rectangles overlap.
    """
    element_ids = [
        get_attr(element, "id")
        for element in elements
        if get_attr(element, "id") in positions
    ]

    occupied = []

    for element_id in element_ids:

        rect = positions[element_id].copy()

        collision = any(
            rectangles_overlap(rect, previous)
            for previous in occupied
        )

        if collision:

            original_x = rect["x"]
            original_y = rect["y"]

            attempt = 0

            while collision and attempt < 100:

                attempt += 1

                rect["x"] = original_x + (
                    attempt % 5
                ) * (NODE_WIDTH + HORIZONTAL_GAP)

                rect["y"] = original_y + (
                    attempt // 5
                ) * (NODE_HEIGHT + VERTICAL_GAP)

                collision = any(
                    rectangles_overlap(rect, previous)
                    for previous in occupied
                )

            positions[element_id] = rect

            REPORT["warnings"].append(
                f"Layout adjustment applied to element: {element_id}"
            )

        occupied.append(rect)


def get_page_dimensions(positions):
    if not positions:
        return 11.0, 8.0

    max_x = max(
        item["x"] + item["width"]
        for item in positions.values()
    )

    max_y = max(
        item["y"] + item["height"]
        for item in positions.values()
    )

    width = max(
        11.0,
        max_x * BPMN_TO_INCH + PAGE_MARGIN
    )

    height = max(
        8.0,
        max_y * BPMN_TO_INCH + PAGE_MARGIN
    )

    return width, height


# ============================================================
# VISIO FORMATTING UTILITIES
# ============================================================

def set_cell(shape, cell_name, formula):
    try:
        shape.CellsU(cell_name).FormulaU = formula
    except Exception:
        pass


def set_shape_text(shape, text):
    """
    Set text on a Visio shape after preparing it for display.

    All text written through this function follows the same
    Persian/Arabic shaping and bidirectional processing path.
    """
    try:
        display_text = prepare_visio_display_text(text)
        shape.Text = display_text
        return True

    except Exception as exc:
        REPORT["warnings"].append(
            f"Could not set shape text: {exc}"
        )
        return False


def format_shape_text(shape, rect, text):
    """
    Configure font, margins, alignment and text wrapping.
    Persian/Arabic text is prepared centrally by set_shape_text.
    """
    text = text or ""

    width = max(rect["width"], 0.1)
    height = max(rect["height"], 0.1)

    # Use the original text length for font-size estimation.
    longest_line = max(
        [len(line) for line in text.splitlines()] or [1]
    )

    estimated_capacity = max(8, int(width * 10))

    font_size = MAX_FONT_SIZE

    if longest_line > estimated_capacity:
        font_size = max(
            MIN_FONT_SIZE,
            MAX_FONT_SIZE * estimated_capacity / longest_line
        )

    if height < 0.45:
        font_size = min(font_size, 8)

    font_size = max(
        MIN_FONT_SIZE,
        min(MAX_FONT_SIZE, font_size)
    )

    set_cell(
        shape,
        "Char.Font",
        f'"{FONT_NAME}"'
    )

    set_cell(
        shape,
        "Char.Size",
        f"{font_size:.1f} pt"
    )

    # Center alignment is retained for BPMN labels.
    set_cell(shape, "Para.HorzAlign", "1")
    set_cell(shape, "Para.VertAlign", "1")

    set_cell(shape, "TxtWrap", "1")

    set_cell(shape, "LeftMargin", "0.06 in")
    set_cell(shape, "RightMargin", "0.06 in")
    set_cell(shape, "TopMargin", "0.04 in")
    set_cell(shape, "BottomMargin", "0.04 in")

    set_shape_text(shape, text)


def format_basic_shape(shape, fill="#FFFFFF", line="#404040"):
    set_cell(shape, "FillForegnd", fill)
    set_cell(shape, "LineColor", line)
    set_cell(shape, "LineWeight", "1.2 pt")
    set_cell(shape, "LinePattern", "1")


def draw_rectangle(page, x1, y1, x2, y2, text="", fill="#FFFFFF"):
    shape = page.DrawRectangle(x1, y1, x2, y2)

    format_basic_shape(shape, fill=fill)

    format_shape_text(
        shape,
        {
            "width": abs(x2 - x1),
            "height": abs(y2 - y1),
        },
        text,
    )

    return shape


def draw_ellipse(page, x1, y1, x2, y2, text="", fill="#FFFFFF"):
    shape = page.DrawOval(x1, y1, x2, y2)

    format_basic_shape(shape, fill=fill)

    format_shape_text(
        shape,
        {
            "width": abs(x2 - x1),
            "height": abs(y2 - y1),
        },
        text,
    )

    return shape


def draw_polygon(page, points, text="", fill="#FFFFFF"):
    """
    Draw a closed polygon using Visio DrawPolyline.
    """
    if not points:
        raise ValueError("Polygon requires at least one point.")

    coordinates = []

    for x, y in points:
        coordinates.extend([
            float(x),
            float(y),
        ])

    coordinates.extend([
        float(points[0][0]),
        float(points[0][1]),
    ])

    shape = page.DrawPolyline(tuple(coordinates), 0)

    format_basic_shape(shape, fill=fill)

    min_x = min(point[0] for point in points)
    max_x = max(point[0] for point in points)
    min_y = min(point[1] for point in points)
    max_y = max(point[1] for point in points)

    format_shape_text(
        shape,
        {
            "width": max_x - min_x,
            "height": max_y - min_y,
        },
        text,
    )

    return shape


def draw_diamond(page, x1, y1, x2, y2, text=""):
    center_x = (x1 + x2) / 2
    center_y = (y1 + y2) / 2

    points = [
        (center_x, y2),
        (x2, center_y),
        (center_x, y1),
        (x1, center_y),
    ]

    return draw_polygon(
        page,
        points,
        text=text,
        fill="#FFF2CC",
    )


def draw_line(page, x1, y1, x2, y2, dashed=False):
    shape = page.DrawLine(x1, y1, x2, y2)

    set_cell(shape, "LineColor", "#404040")
    set_cell(shape, "LineWeight", "1.1 pt")

    if dashed:
        set_cell(shape, "LinePattern", "2")
    else:
        set_cell(shape, "LinePattern", "1")

    return shape


# ============================================================
# BPMN SHAPE RENDERING
# ============================================================

def bpmn_to_visio_coordinates(rect, page_height):
    """
    BPMN uses top-left origin.
    Visio uses bottom-left origin.
    """
    x1 = rect["x"] * BPMN_TO_INCH
    x2 = (
        rect["x"] + rect["width"]
    ) * BPMN_TO_INCH

    y_top = page_height - (
        rect["y"] * BPMN_TO_INCH
    )

    y_bottom = page_height - (
        (rect["y"] + rect["height"]) * BPMN_TO_INCH
    )

    return x1, y_bottom, x2, y_top


def render_event(page, element, rect, page_height):
    x1, y1, x2, y2 = bpmn_to_visio_coordinates(
        rect,
        page_height,
    )

    event_type = local_name(element.tag)
    definitions = get_event_definition(element)

    marker = ""

    if definitions:
        marker = {
            "Message": "✉",
            "Timer": "◷",
            "Error": "!",
            "Signal": "⚑",
            "Escalation": "↑",
            "Conditional": "?",
            "Cancel": "×",
            "Compensate": "↶",
            "Terminate": "■",
            "Link": "↗",
        }.get(definitions[0], "")

    name = get_element_name(element)

    label = marker

    if name:
        label = (
            name
            if not marker
            else marker + "\n" + name
        )

    shape = draw_ellipse(
        page,
        x1,
        y1,
        x2,
        y2,
        label,
        fill="#FFFFFF",
    )

    if event_type == "startEvent":
        set_cell(shape, "LineWeight", "1.5 pt")

    elif event_type == "endEvent":
        set_cell(shape, "LineWeight", "3 pt")

    elif event_type == "boundaryEvent":
        set_cell(shape, "LinePattern", "2")

    return shape


def render_gateway(page, element, rect, page_height):
    x1, y1, x2, y2 = bpmn_to_visio_coordinates(
        rect,
        page_height,
    )

    marker = get_gateway_marker(element)
    name = get_element_name(element)

    label = marker

    if name:
        label = (
            marker + "\n" + name
            if marker
            else name
        )

    return draw_diamond(
        page,
        x1,
        y1,
        x2,
        y2,
        label,
    )


def render_activity(page, element, rect, page_height):
    x1, y1, x2, y2 = bpmn_to_visio_coordinates(
        rect,
        page_height,
    )

    element_type = local_name(element.tag)
    name = get_element_name(element)

    marker = {
        "userTask": "User",
        "manualTask": "Manual",
        "serviceTask": "Service",
        "scriptTask": "Script",
        "businessRuleTask": "Rule",
        "sendTask": "Send",
        "receiveTask": "Receive",
        "callActivity": "Call",
        "subProcess": "SubProcess",
        "transaction": "Transaction",
        "choreographyTask": "Choreography",
    }.get(element_type, "")

    label = name

    if marker:
        label = (
            f"{marker}\n{name}"
            if name
            else marker
        )

    fill = "#FFFFFF"

    if element_type in {
        "callActivity",
        "subProcess",
        "transaction",
    }:
        fill = "#EAF2F8"

    shape = draw_rectangle(
        page,
        x1,
        y1,
        x2,
        y2,
        label,
        fill=fill,
    )

    if element_type in {
        "subProcess",
        "transaction",
    }:
        set_cell(shape, "LineWeight", "1.8 pt")

    return shape


def render_data_element(page, element, rect, page_height):
    x1, y1, x2, y2 = bpmn_to_visio_coordinates(
        rect,
        page_height,
    )

    name = get_element_name(element)
    element_type = local_name(element.tag)

    if "Store" in element_type:

        shape = draw_rectangle(
            page,
            x1,
            y1,
            x2,
            y2,
            name,
            fill="#E2F0D9",
        )

        set_cell(shape, "LineWeight", "2 pt")

    else:

        shape = draw_rectangle(
            page,
            x1,
            y1,
            x2,
            y2,
            name,
            fill="#F2F2F2",
        )

    return shape


def render_annotation(page, element, rect, page_height):
    x1, y1, x2, y2 = bpmn_to_visio_coordinates(
        rect,
        page_height,
    )

    text = ""

    for child in element:
        if local_name(child.tag) == "text":
            text = child.text or ""
            break

    shape = draw_rectangle(
        page,
        x1,
        y1,
        x2,
        y2,
        text,
        fill="#FFF9E6",
    )

    set_cell(shape, "LinePattern", "2")

    return shape


def render_container(page, element, rect, page_height):
    x1, y1, x2, y2 = bpmn_to_visio_coordinates(
        rect,
        page_height,
    )

    name = get_element_name(element)
    element_type = local_name(element.tag)

    if element_type == "participant":
        fill = "#F2F2F2"
    else:
        fill = "#FFFFFF"

    shape = draw_rectangle(
        page,
        x1,
        y1,
        x2,
        y2,
        name,
        fill=fill,
    )

    set_cell(shape, "LineWeight", "1.5 pt")

    return shape


def render_bpmn_element(page, element, rect, page_height):
    element_type = local_name(element.tag)

    if element_type in EVENT_TYPES:
        return render_event(
            page, element, rect, page_height
        )

    if element_type in GATEWAY_TYPES:
        return render_gateway(
            page, element, rect, page_height
        )

    if element_type in TASK_TYPES:
        return render_activity(
            page, element, rect, page_height
        )

    if is_data_element(element):
        return render_data_element(
            page, element, rect, page_height
        )

    if is_annotation(element):
        return render_annotation(
            page, element, rect, page_height
        )

    if is_container(element):
        return render_container(
            page, element, rect, page_height
        )

    return None


# ============================================================
# BPMN FLOW RENDERING
# ============================================================

def get_flow_endpoints(element):
    source = get_attr(element, "sourceRef")
    target = get_attr(element, "targetRef")

    return source, target


def get_node_center(rect, page_height):
    x = (
        rect["x"] + rect["width"] / 2
    ) * BPMN_TO_INCH

    y = page_height - (
        rect["y"] + rect["height"] / 2
    ) * BPMN_TO_INCH

    return x, y


def render_flow(page, flow, positions, page_height, edge_information):
    flow_type = local_name(flow.tag)

    source_id, target_id = get_flow_endpoints(flow)

    if not source_id or not target_id:
        REPORT["warnings"].append(
            f"Flow without complete source/target: "
            f"{get_attr(flow, 'id')}"
        )
        return

    if (
        source_id not in positions
        or target_id not in positions
    ):
        REPORT["warnings"].append(
            f"Flow endpoints not found: "
            f"{get_attr(flow, 'id')}"
        )
        return

    points = edge_information.get(
        get_attr(flow, "id"),
        [],
    )

    if len(points) >= 2:

        converted_points = []

        for point in points:

            x = point["x"] * BPMN_TO_INCH

            y = page_height - (
                point["y"] * BPMN_TO_INCH
            )

            converted_points.append((x, y))

    else:

        start = get_node_center(
            positions[source_id],
            page_height,
        )

        end = get_node_center(
            positions[target_id],
            page_height,
        )

        converted_points = [
            start,
            end,
        ]

    dashed = flow_type in {
        "messageFlow",
        "association",
        "dataInputAssociation",
        "dataOutputAssociation",
        "conversationLink",
    }

    for index in range(
        len(converted_points) - 1
    ):

        p1 = converted_points[index]
        p2 = converted_points[index + 1]

        shape = draw_line(
            page,
            p1[0],
            p1[1],
            p2[0],
            p2[1],
            dashed=dashed,
        )

        if flow_type == "sequenceFlow":
            set_cell(shape, "EndArrow", "4")

        elif flow_type == "messageFlow":
            set_cell(shape, "EndArrow", "4")
            set_cell(shape, "LinePattern", "2")

        REPORT["converted_flows"] += 1


# ============================================================
# PROCESS AND DIAGRAM EXTRACTION
# ============================================================

def collect_renderable_elements(root):
    """Collect all diagrammable BPMN elements."""
    elements = []

    accepted = (
        FLOW_NODE_TYPES
        | DATA_TYPES
        | ANNOTATION_TYPES
        | {"participant", "lane"}
    )

    for element in root.iter():

        if local_name(element.tag) in accepted:
            if get_attr(element, "id"):
                elements.append(element)

    return elements


def collect_flows(root):
    return [
        element
        for element in root.iter()
        if local_name(element.tag) in FLOW_TYPES
    ]


def get_diagram_elements(root, diagram):
    """
    Return semantic elements belonging to a BPMN diagram plane.
    """
    element_index = get_document_elements(root)

    plane_id = diagram.get("plane_element")

    if (
        not plane_id
        or plane_id not in element_index
    ):
        return (
            collect_renderable_elements(root),
            collect_flows(root),
        )

    plane_element = element_index[plane_id]
    included_ids = set()

    for item in plane_element.iter():

        item_id = get_attr(item, "id")

        if item_id:
            included_ids.add(item_id)

        for attribute in (
            "processRef",
            "participantRef",
        ):
            ref = get_attr(item, attribute)

            if ref:
                included_ids.add(ref)

    # Participants may refer to processes outside the collaboration.
    for item in plane_element.iter():

        process_ref = get_attr(
            item,
            "processRef",
        )

        if (
            process_ref
            and process_ref in element_index
        ):

            for child in element_index[process_ref].iter():

                child_id = get_attr(child, "id")

                if child_id:
                    included_ids.add(child_id)

    # Expand descendants recursively.
    changed = True

    while changed:
        changed = False

        for element_id in list(included_ids):

            element = element_index.get(element_id)

            if element is None:
                continue

            for child in element.iter():

                child_id = get_attr(child, "id")

                if (
                    child_id
                    and child_id not in included_ids
                ):
                    included_ids.add(child_id)
                    changed = True

    accepted = (
        FLOW_NODE_TYPES
        | DATA_TYPES
        | ANNOTATION_TYPES
        | {"participant", "lane"}
    )

    elements = [
        element
        for element in root.iter()
        if (
            get_attr(element, "id") in included_ids
            and local_name(element.tag) in accepted
        )
    ]

    flows = [
        element
        for element in root.iter()
        if (
            get_attr(element, "id") in included_ids
            and local_name(element.tag) in FLOW_TYPES
        )
    ]

    if not elements:
        elements = collect_renderable_elements(root)
        flows = collect_flows(root)

    return elements, flows


# ============================================================
# VISIO DOCUMENT CREATION
# ============================================================

def create_visio_application():
    app = win32com.client.DispatchEx(
        "Visio.Application"
    )

    app.Visible = True

    return app


def configure_page(page, width, height):
    """Configure Visio page size."""
    try:
        page.PageSheet.CellsU(
            "PageWidth"
        ).FormulaU = f"{width:.4f} in"

        page.PageSheet.CellsU(
            "PageHeight"
        ).FormulaU = f"{height:.4f} in"

    except Exception:
        REPORT["warnings"].append(
            "Could not configure the exact Visio page dimensions."
        )


def add_report_page(document):
    """Add a final page containing the conversion report."""
    try:
        page = document.Pages.Add()
        page.Name = "Conversion Report"

        draw_rectangle(
            page,
            0.5,
            0.5,
            10.5,
            7.5,
            "BPMN 2.0 Conversion Report",
            fill="#EAF2F8",
        )

        y = 6.8

        lines = [
            f"Converter version: {VERSION}",
            f"Converted elements: {REPORT['converted_elements']}",
            f"Converted flow segments: {REPORT['converted_flows']}",
            f"Pages: {REPORT['pages']}",
            "",
            "Warnings:",
        ]

        lines.extend(
            REPORT["warnings"][:35]
        )

        if REPORT["errors"]:
            lines.append("")
            lines.append("Errors:")
            lines.extend(REPORT["errors"][:20])

        for line in lines:

            draw_rectangle(
                page,
                0.7,
                y - 0.25,
                10.3,
                y,
                line,
                fill="#FFFFFF",
            )

            y -= 0.32

            if y < 0.6:
                break

    except Exception as exc:
        REPORT["warnings"].append(
            f"Could not create report page: {exc}"
        )


# ============================================================
# MAIN CONVERSION FUNCTION
# ============================================================

def convert_bpmn_to_visio(source_file, output_file=None):
    source_file = Path(source_file)

    if not source_file.exists():
        raise FileNotFoundError(
            f"Input file does not exist: {source_file}"
        )

    if output_file is None:
        output_file = source_file.with_suffix(".vsdx")

    output_file = Path(output_file)

    output_file.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    # Reset report for each conversion.
    REPORT["warnings"].clear()
    REPORT["errors"].clear()
    REPORT["converted_elements"] = 0
    REPORT["converted_flows"] = 0
    REPORT["pages"] = 0

    print("\nReading BPMN XML...")

    try:
        tree = ET.parse(source_file)
        root = tree.getroot()

    except ET.ParseError as exc:
        raise ValueError(
            f"Invalid XML/BPMN file: {exc}"
        )

    if local_name(root.tag) != "definitions":
        raise ValueError(
            "The input XML root element is not BPMN definitions."
        )

    diagrams, global_shapes, global_edges = (
        get_diagram_information(root)
    )

    if not diagrams:

        REPORT["warnings"].append(
            "No BPMN-DI diagram was found. "
            "Automatic layout will be used."
        )

        diagrams = [{
            "shapes": {},
            "edges": {},
            "plane_element": "",
        }]

    print(
        f"Found {len(diagrams)} BPMN diagram(s)."
    )

    app = None
    document = None

    try:
        print("Starting Microsoft Visio...")

        app = create_visio_application()
        document = app.Documents.Add("")

        # Use the first page created by Visio.
        first_page = document.Pages.Item(1)

        for diagram_index, diagram in enumerate(diagrams):

            if diagram_index == 0:
                page = first_page
            else:
                page = document.Pages.Add()

            page.Name = (
                f"BPMN Diagram {diagram_index + 1}"
            )

            elements, flows = get_diagram_elements(
                root,
                diagram,
            )

            positions = {}

            # Load original BPMN-DI positions.
            for element in elements:

                element_id = get_attr(
                    element,
                    "id",
                )

                if element_id in diagram["shapes"]:
                    positions[element_id] = (
                        diagram["shapes"][element_id].copy()
                    )

            # Generate positions for elements without BPMN-DI.
            missing_elements = [
                element
                for element in elements
                if (
                    get_attr(element, "id") not in positions
                    and (
                        is_flow_node(element)
                        or is_data_element(element)
                        or is_annotation(element)
                    )
                )
            ]

            automatic_positions = auto_layout(
                missing_elements
            )

            positions.update(automatic_positions)

            correct_overlaps(
                positions,
                missing_elements,
            )

            page_width, page_height = (
                get_page_dimensions(positions)
            )

            configure_page(
                page,
                page_width,
                page_height,
            )

            # ------------------------------------------------
            # 1. Draw participants and lanes first.
            # ------------------------------------------------

            containers = [
                element
                for element in elements
                if is_container(element)
            ]

            for element in containers:

                element_id = get_attr(
                    element,
                    "id",
                )

                rect = positions.get(element_id)

                if rect is None:
                    continue

                try:
                    render_container(
                        page,
                        element,
                        rect,
                        page_height,
                    )

                    REPORT["converted_elements"] += 1

                except Exception as exc:
                    REPORT["warnings"].append(
                        f"Container drawing failed "
                        f"({element_id}): {exc}"
                    )

            # ------------------------------------------------
            # 2. Draw flows behind nodes.
            # ------------------------------------------------

            for flow in flows:

                try:
                    render_flow(
                        page,
                        flow,
                        positions,
                        page_height,
                        diagram["edges"],
                    )

                except Exception as exc:
                    REPORT["warnings"].append(
                        f"Flow drawing failed "
                        f"({get_attr(flow, 'id')}): {exc}"
                    )

            # ------------------------------------------------
            # 3. Draw activities, events, gateways and data.
            # ------------------------------------------------

            for element in elements:

                if is_container(element):
                    continue

                element_id = get_attr(
                    element,
                    "id",
                )

                rect = positions.get(element_id)

                if rect is None:
                    continue

                try:
                    shape = render_bpmn_element(
                        page,
                        element,
                        rect,
                        page_height,
                    )

                    if shape is not None:
                        REPORT["converted_elements"] += 1

                except Exception as exc:
                    REPORT["warnings"].append(
                        f"Element drawing failed "
                        f"({element_id}): {exc}"
                    )

            REPORT["pages"] += 1

        # Add a separate conversion report page.
        add_report_page(document)

        print("Saving Visio document...")

        document.SaveAs(
            str(output_file.resolve())
        )

        print("\nCONVERSION COMPLETED")
        print(f"Output: {output_file}")

        return output_file, REPORT

    except Exception:

        REPORT["errors"].append(
            traceback.format_exc()
        )

        raise

    finally:
        # Keep Visio application and document open.
        pass


# ============================================================
# COMMAND-LINE USER INTERFACE
# ============================================================

def choose_input_file():
    import tkinter as tk
    from tkinter import filedialog
    from tkinter import messagebox

    window = None

    try:
        window = tk.Tk()
        window.withdraw()

        window.update_idletasks()
        window.attributes("-topmost", True)
        window.lift()
        window.focus_force()

        filename = filedialog.askopenfilename(
            parent=window,
            title="Select BPMN 2.0 File",
            filetypes=[
                ("BPMN 2.0 Files", "*.bpmn"),
                ("XML Files", "*.xml"),
                ("All Files", "*.*"),
            ],
            defaultextension=".bpmn",
        )

        return filename if filename else ""

    except Exception as exc:
        print(f"File selection error: {exc}")

        try:
            messagebox.showerror(
                "File Selection Error",
                str(exc),
                parent=window,
            )
        except Exception:
            pass

        return ""

    finally:
        if window is not None:
            try:
                window.destroy()
            except Exception:
                pass


def choose_output_file(default_path):
    window = None

    try:
        import tkinter as tk
        from tkinter import filedialog

        window = tk.Tk()
        window.withdraw()
        window.attributes("-topmost", True)

        filename = filedialog.asksaveasfilename(
            parent=window,
            title="Save Visio File",
            initialfile=Path(default_path).name,
            defaultextension=".vsdx",
            filetypes=[
                ("Visio Drawing", "*.vsdx"),
            ],
        )

        return filename

    except Exception:
        return ""

    finally:
        if window is not None:
            try:
                window.destroy()
            except Exception:
                pass


def print_report(report):
    print("\n" + "=" * 65)
    print("CONVERSION REPORT")
    print("=" * 65)

    print(
        f"Converted elements : "
        f"{report['converted_elements']}"
    )

    print(
        f"Converted flow segments : "
        f"{report['converted_flows']}"
    )

    print(
        f"Created pages : "
        f"{report['pages']}"
    )

    print(
        f"Warnings : "
        f"{len(report['warnings'])}"
    )

    for warning in report["warnings"]:
        print("WARNING:", warning)

    if report["errors"]:
        print("\nERRORS:")

        for error in report["errors"]:
            print(error)

    print("=" * 65)


def main():
    print("\n" + "=" * 65)
    print(APP_NAME)
    print(f"Version {VERSION}")
    print("=" * 65)

    print("\n1. Select BPMN file")
    print("2. Exit")

    choice = input(
        "\nSelect option: "
    ).strip()

    if choice == "2":
        return

    if choice != "1":
        print("Invalid option.")
        return

    source = choose_input_file()

    if not source:
        source = input(
            "\nEnter BPMN file path: "
        ).strip().strip('"')

    if not source:
        print("No input file selected.")
        return

    default_output = str(
        Path(source).with_suffix(".vsdx")
    )

    output = choose_output_file(
        default_output
    )

    if not output:
        output = input(
            "\nEnter output VSDX path "
            "(press ENTER for default): "
        ).strip().strip('"')

    if not output:
        output = default_output

    try:
        result, report = convert_bpmn_to_visio(
            source,
            output,
        )

        print_report(report)

        print("\nVisio file created:")
        print(result)

    except Exception as exc:

        print("\nCONVERSION FAILED")
        print(str(exc))

        print("\nDetailed error:")
        traceback.print_exc()

    input("\nPress ENTER to exit...")


if __name__ == "__main__":
    main()
