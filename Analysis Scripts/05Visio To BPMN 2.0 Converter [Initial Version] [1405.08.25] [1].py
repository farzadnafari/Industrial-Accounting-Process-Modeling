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
VISIO TO BPMN 2.0 CONVERTER
Version: 1.0

Input:
    Microsoft Visio .vsdx

Output:
    BPMN 2.0 XML (.bpmn)
    Conversion report (.txt)

Requirements:
    Windows
    Microsoft Visio Desktop
    Python 3.10+
    pywin32

The user only needs to draw shapes, connectors and labels in Visio.
No BPMN Shape Data or manual BPMN configuration is required.

Automatic mapping:
    Circle / Ellipse       -> BPMN Event
    Diamond / Rhombus      -> Exclusive Gateway
    Rounded Rectangle      -> Task
    Rectangle              -> Task
    Solid Connector        -> Sequence Flow
    Dashed Connector       -> Association

Important:
    This is a heuristic converter. A plain Visio drawing does not contain
    enough information to infer every BPMN semantic distinction.
    Ambiguous cases are handled conservatively and recorded in the report.
"""

from __future__ import annotations

import re
import sys
import traceback
import unicodedata
import xml.etree.ElementTree as ET

from pathlib import Path
from collections import defaultdict

import win32com.client


# ============================================================
# CONFIGURATION
# ============================================================

APP_NAME = "Visio to BPMN 2.0 Converter"
VERSION = "1.0"

BPMN_NS = "http://www.omg.org/spec/BPMN/20100524/MODEL"
BPMNDI_NS = "http://www.omg.org/spec/BPMN/20100524/DI"
DC_NS = "http://www.omg.org/spec/DD/20100524/DC"
DI_NS = "http://www.omg.org/spec/DD/20100524/DI"

ET.register_namespace("bpmn", BPMN_NS)
ET.register_namespace("bpmndi", BPMNDI_NS)
ET.register_namespace("dc", DC_NS)
ET.register_namespace("di", DI_NS)

NS = {
    "bpmn": BPMN_NS,
    "bpmndi": BPMNDI_NS,
    "dc": DC_NS,
    "di": DI_NS,
}

# Visio dimensions are in inches.
# BPMN-DI dimensions are conventionally represented in diagram units.
INCH_TO_BPMN = 96.0

# Automatic layout adjustment for overlapping or missing shapes.
MIN_NODE_WIDTH = 60.0
MIN_NODE_HEIGHT = 40.0
DEFAULT_NODE_WIDTH = 120.0
DEFAULT_NODE_HEIGHT = 70.0

# If a connector is not glued, this is the maximum distance
# used to associate its endpoint with a nearby shape.
ENDPOINT_TOLERANCE = 0.35

REPORT = {
    "warnings": [],
    "errors": [],
    "corrections": [],
    "elements": 0,
    "flows": 0,
}


def qname(namespace, name):
    return f"{{{namespace}}}{name}"


def log_warning(message):
    REPORT["warnings"].append(str(message))


def log_correction(message):
    REPORT["corrections"].append(str(message))


def safe_id(value, prefix="Element"):
    """
    Generate a valid and unique XML-friendly ID.
    """
    value = str(value or "").strip()

    value = re.sub(
        r"[^a-zA-Z0-9_.-]",
        "_",
        value,
    )

    if not value:
        value = prefix

    if not re.match(r"^[a-zA-Z_]", value):
        value = prefix + "_" + value

    return value


def clean_text(value):
    if value is None:
        return ""

    value = str(value)

    value = value.replace("\x00", "")
    value = value.replace("\r\n", "\n")
    value = value.replace("\r", "\n")

    return value.strip()


# ============================================================
# VISIO COM UTILITIES
# ============================================================

def get_cell_result(shape, cell_name, default=0.0):
    try:
        return float(
            shape.CellsU(cell_name).ResultIU
        )
    except Exception:
        return default


def get_shape_text(shape):
    try:
        return clean_text(shape.Text)
    except Exception:
        return ""


def get_shape_name(shape):
    try:
        return str(shape.NameU)
    except Exception:
        return ""


def get_master_name(shape):
    try:
        return str(shape.Master.NameU)
    except Exception:
        return ""


def get_shape_metadata(shape):
    """
    Read basic geometric and naming information from a Visio shape.
    """
    width = get_cell_result(shape, "Width", 1.0)
    height = get_cell_result(shape, "Height", 0.6)

    pin_x = get_cell_result(shape, "PinX", 0.0)
    pin_y = get_cell_result(shape, "PinY", 0.0)

    left = pin_x - width / 2
    bottom = pin_y - height / 2

    right = pin_x + width / 2
    top = pin_y + height / 2

    return {
        "shape": shape,
        "name": get_shape_name(shape),
        "master": get_master_name(shape),
        "text": get_shape_text(shape),
        "width": width,
        "height": height,
        "center_x": pin_x,
        "center_y": pin_y,
        "left": left,
        "right": right,
        "bottom": bottom,
        "top": top,
        "one_d": is_one_d_shape(shape),
    }


def is_one_d_shape(shape):
    try:
        return bool(shape.OneD)
    except Exception:
        return False


def get_page_size(page):
    width = get_cell_result(
        page.PageSheet,
        "PageWidth",
        11.0,
    )

    height = get_cell_result(
        page.PageSheet,
        "PageHeight",
        8.5,
    )

    return width, height


def get_page_shapes(page):
    """
    Collect shapes from a page, including children of groups.
    """
    result = []

    def walk(collection):
        try:
            count = collection.Count
        except Exception:
            return

        for index in range(1, count + 1):
            try:
                shape = collection.Item(index)
            except Exception:
                continue

            result.append(shape)

            try:
                if shape.Type == 1 and shape.Shapes.Count:
                    walk(shape.Shapes)
            except Exception:
                pass

    walk(page.Shapes)

    return result


# ============================================================
# VISIO SHAPE CLASSIFICATION
# ============================================================

def classify_shape(metadata):
    """
    Determine the most likely BPMN element from Visio geometry
    and shape/master names.

    This deliberately does not infer advanced task or event types
    from arbitrary text.
    """
    name = metadata["name"].lower()
    master = metadata["master"].lower()

    combined = f"{name} {master}"

    width = abs(metadata["width"])
    height = abs(metadata["height"])

    if metadata["one_d"]:
        return "connector"

    # Explicit naming is more reliable than geometry.
    if any(word in combined for word in (
        "diamond",
        "rhombus",
        "gateway",
        "decision",
    )):
        return "gateway"

    if any(word in combined for word in (
        "ellipse",
        "oval",
        "circle",
        "event",
    )):
        return "event"

    if any(word in combined for word in (
        "round",
        "rounded",
        "task",
        "activity",
    )):
        return "task"

    if any(word in combined for word in (
        "annotation",
        "note",
        "comment",
    )):
        return "annotation"

    # Basic geometry heuristics.
    if width > 0 and height > 0:

        ratio = width / height

        # Circles and ellipses.
        if 0.78 <= ratio <= 1.22:
            if width <= 1.25 and height <= 1.25:
                return "event"

        # Diamond detection is less reliable for plain Visio shapes.
        # Master/name detection above is preferred.
        if 0.75 <= ratio <= 1.25:
            if width <= 2.2 and height <= 2.2:
                if "decision" in combined:
                    return "gateway"

    # Default: a plain rectangle becomes a Task.
    return "task"


# ============================================================
# VISIO CONNECTOR ANALYSIS
# ============================================================

def get_connector_text(shape):
    return get_shape_text(shape)


def get_connector_line_pattern(shape):
    try:
        return int(
            shape.CellsU("LinePattern").ResultIU
        )
    except Exception:
        return 1


def get_connector_direction(shape):
    """
    Detect whether the visible arrowhead is at the beginning
    or the end of a Visio connector.

    Returns:
        True  = Reverse direction (End -> Begin)
        False = Normal direction (Begin -> End)
    """
    begin_arrow = 0
    end_arrow = 0

    try:
        begin_arrow = int(
            shape.CellsU("BeginArrow").ResultIU
        )
    except Exception:
        pass

    try:
        end_arrow = int(
            shape.CellsU("EndArrow").ResultIU
        )
    except Exception:
        pass

    if begin_arrow != 0 and end_arrow == 0:
        return True

    return False


def is_dashed_connector(shape):
    return get_connector_line_pattern(shape) != 1


def get_connector_endpoints(shape):
    """
    Return connector endpoints in Visio page coordinates.
    """
    begin_x = get_cell_result(shape, "BeginX", 0.0)
    begin_y = get_cell_result(shape, "BeginY", 0.0)

    end_x = get_cell_result(shape, "EndX", 0.0)
    end_y = get_cell_result(shape, "EndY", 0.0)

    return (
        (begin_x, begin_y),
        (end_x, end_y),
    )


def get_glued_endpoints(shape):
    """
    Extract glued shapes from Visio Connects.

    A Visio connector normally has connections to its
    BeginX and EndX cells when glued to other shapes.
    """
    begin_shape = None
    end_shape = None

    try:
        connects = shape.Connects

        for index in range(1, connects.Count + 1):

            connection = connects.Item(index)

            try:
                from_cell = str(
                    connection.FromCell.NameU
                ).lower()

                target = connection.ToSheet

                if from_cell == "beginx":
                    begin_shape = target

                elif from_cell == "endx":
                    end_shape = target

            except Exception:
                continue

    except Exception:
        pass

    return begin_shape, end_shape


def find_nearest_node(point, nodes, tolerance=ENDPOINT_TOLERANCE):
    """
    Find the nearest node using distance to its bounding rectangle,
    rather than distance to its center.
    """
    x, y = point

    best_node = None
    best_distance = float("inf")

    for node in nodes:

        left = min(node["left"], node["right"])
        right = max(node["left"], node["right"])
        bottom = min(node["bottom"], node["top"])
        top = max(node["bottom"], node["top"])

        dx = max(left - x, 0.0, x - right)
        dy = max(bottom - y, 0.0, y - top)

        distance = (dx * dx + dy * dy) ** 0.5

        if distance < best_distance:
            best_distance = distance
            best_node = node

    if best_node is not None and best_distance <= tolerance:
        return best_node

    return None


# ============================================================
# BPMN XML ELEMENT CREATION
# ============================================================

def create_definitions():
    root = ET.Element(
        qname(BPMN_NS, "definitions"),
        {
            "id": "Definitions_Visio_Conversion",
            "targetNamespace": "http://example.com/bpmn",
            "exporter": APP_NAME,
            "exporterVersion": VERSION,
        },
    )

    return root


def add_process(root, process_id, process_name):
    process = ET.SubElement(
        root,
        qname(BPMN_NS, "process"),
        {
            "id": process_id,
            "name": process_name,
            "isExecutable": "false",
        },
    )

    return process


def add_bpmn_node(process, node_type, node_id, name):
    attributes = {
        "id": node_id,
    }

    if name:
        attributes["name"] = name

    return ET.SubElement(
        process,
        qname(BPMN_NS, node_type),
        attributes,
    )


def add_sequence_flow(process, flow_id, source_id, target_id, name=""):
    attributes = {
        "id": flow_id,
        "sourceRef": source_id,
        "targetRef": target_id,
    }

    if name:
        attributes["name"] = name

    return ET.SubElement(
        process,
        qname(BPMN_NS, "sequenceFlow"),
        attributes,
    )


# ============================================================
# BPMN-DI CREATION
# ============================================================

def create_diagram(root, process_id):
    diagram = ET.SubElement(
        root,
        qname(BPMNDI_NS, "BPMNDiagram"),
        {
            "id": "BPMNDiagram_1",
        },
    )

    plane = ET.SubElement(
        diagram,
        qname(BPMNDI_NS, "BPMNPlane"),
        {
            "id": "BPMNPlane_1",
            "bpmnElement": process_id,
        },
    )

    return plane


def add_di_shape(plane, bpmn_id, metadata, page_height):
    """
    Create BPMN-DI shape bounds with automatic text-aware sizing.

    The original shape center is preserved.
    Width and height are increased when necessary to accommodate text.
    """

    text = clean_text(metadata.get("text", ""))

    original_width = abs(metadata["width"]) * INCH_TO_BPMN
    original_height = abs(metadata["height"]) * INCH_TO_BPMN

    center_x = metadata["center_x"] * INCH_TO_BPMN

    center_y = (
        page_height - metadata["center_y"]
    ) * INCH_TO_BPMN

    # Minimum dimensions.
    width = max(original_width, MIN_NODE_WIDTH)
    height = max(original_height, MIN_NODE_HEIGHT)

    # Estimate required text space.
    if text:

        lines = text.split("\n")

        longest_line = max(
            len(line)
            for line in lines
        )

        line_count = len(lines)

        # Approximate text width for a readable BPMN label.
        estimated_text_width = (
            longest_line * 7.5 + 36
        )

        estimated_text_height = (
            line_count * 20 + 28
        )

        width = max(
            width,
            estimated_text_width,
        )

        height = max(
            height,
            estimated_text_height,
        )

    # Additional safety margin.
    width += 8
    height += 8

    # Preserve the original center.
    x = center_x - width / 2
    y = center_y - height / 2

    # Store the actual generated bounds for connector calculations.
    metadata["di_bounds"] = {
        "x": x,
        "y": y,
        "width": width,
        "height": height,
        "center_x": center_x,
        "center_y": center_y,
    }

    shape = ET.SubElement(
        plane,
        qname(BPMNDI_NS, "BPMNShape"),
        {
            "id": f"DI_{bpmn_id}",
            "bpmnElement": bpmn_id,
        },
    )

    ET.SubElement(
        shape,
        qname(DC_NS, "Bounds"),
        {
            "x": f"{x:.3f}",
            "y": f"{y:.3f}",
            "width": f"{width:.3f}",
            "height": f"{height:.3f}",
        },
    )

def get_di_connection_point(node, other_node):
    """
    Calculate the intersection of the center-to-center line
    with the actual BPMN-DI shape boundary.
    """

    bounds = node.get("di_bounds")

    if not bounds:
        x = node["center_x"] * INCH_TO_BPMN
        y = node["center_y"] * INCH_TO_BPMN

        return x, y

    cx = bounds["center_x"]
    cy = bounds["center_y"]

    other_bounds = other_node.get("di_bounds")

    if other_bounds:
        ox = other_bounds["center_x"]
        oy = other_bounds["center_y"]
    else:
        ox = other_node["center_x"] * INCH_TO_BPMN
        oy = other_node["center_y"] * INCH_TO_BPMN

    dx = ox - cx
    dy = oy - cy

    if abs(dx) < 0.0001 and abs(dy) < 0.0001:
        return cx, cy

    half_width = bounds["width"] / 2
    half_height = bounds["height"] / 2

    kind = node.get("kind", "task")

    if kind == "event":
        # Ellipse boundary.
        scale = 1.0 / (
            (
                (dx / half_width) ** 2
                + (dy / half_height) ** 2
            ) ** 0.5
        )

    elif kind == "gateway":
        # Diamond boundary.
        denominator = (
            abs(dx) / half_width
            + abs(dy) / half_height
        )

        scale = 1.0 / denominator if denominator else 1.0

    else:
        # Rectangle boundary.
        scale_x = (
            half_width / abs(dx)
            if abs(dx) > 0.0001
            else float("inf")
        )

        scale_y = (
            half_height / abs(dy)
            if abs(dy) > 0.0001
            else float("inf")
        )

        scale = min(scale_x, scale_y)

    return (
        cx + dx * scale,
        cy + dy * scale,
    )


def get_flow_di_endpoints(source_node, target_node):
    """
    Calculate flow waypoints at the boundaries of the source
    and target BPMN shapes.
    """

    start = get_di_connection_point(
        source_node,
        target_node,
    )

    end = get_di_connection_point(
        target_node,
        source_node,
    )

    return start, end


def add_di_edge(plane, flow_id, source_node, target_node):
    """
    Create BPMN-DI edge with correctly directed boundary waypoints.
    """

    start, end = get_flow_di_endpoints(
        source_node,
        target_node,
    )

    edge = ET.SubElement(
        plane,
        qname(BPMNDI_NS, "BPMNEdge"),
        {
            "id": f"DI_{flow_id}",
            "bpmnElement": flow_id,
        },
    )

    ET.SubElement(
        edge,
        qname(DI_NS, "waypoint"),
        {
            "x": f"{start[0]:.3f}",
            "y": f"{start[1]:.3f}",
        },
    )

    ET.SubElement(
        edge,
        qname(DI_NS, "waypoint"),
        {
            "x": f"{end[0]:.3f}",
            "y": f"{end[1]:.3f}",
        },
    )


# ============================================================
# AUTOMATIC BPMN SEMANTIC CORRECTIONS
# ============================================================

def determine_event_type(metadata, incoming_count, outgoing_count):
    """
    Advanced BPMN 2.0 Event Type Detection Engine.

    Detection priority:
        1. Explicit BPMN metadata
        2. Explicit event terminology in shape text/name
        3. Event definition keywords
        4. Graph position
        5. Conservative fallback

    Supported event categories:
        - Start Events
        - End Events
        - Intermediate Catch Events
        - Intermediate Throw Events
        - Boundary Events (when metadata explicitly identifies them)

    Supports Persian and English terminology.

    Important:
        This function classifies event types only.
        It does not determine the detailed BPMN event definition
        such as Timer, Message, Signal, Error, Escalation, etc.
    """

    # ---------------------------------------------------------
    # 1. TEXT NORMALIZATION
    # ---------------------------------------------------------

    def normalize_text(value):
        if value is None:
            return ""

        text = str(value)

        # Unicode normalization
        text = unicodedata.normalize("NFKC", text)

        # Persian / Arabic character normalization
        replacements = {
            "ي": "ی",
            "ى": "ی",
            "ك": "ک",
            "ۀ": "ه",
            "ة": "ه",
            "\u0640": "",
        }

        for old, new in replacements.items():
            text = text.replace(old, new)

        # Normalize whitespace
        text = re.sub(r"\s+", " ", text)

        return text.strip().lower()

    # ---------------------------------------------------------
    # 2. SAFE METADATA ACCESS
    # ---------------------------------------------------------

    metadata = metadata or {}

    def get_metadata_value(*keys):
        for key in keys:
            value = metadata.get(key)

            if value is not None and str(value).strip():
                return value

        return ""

    shape_text = normalize_text(
        get_metadata_value(
            "text",
            "shape_text",
            "label",
            "display_text",
            "name",
            "shape_name",
        )
    )

    shape_name = normalize_text(
        get_metadata_value(
            "shape_name",
            "name",
            "master_name",
            "master",
        )
    )

    shape_type = normalize_text(
        get_metadata_value(
            "bpmn_type",
            "element_type",
            "shape_type",
            "type",
        )
    )

    event_subtype = normalize_text(
        get_metadata_value(
            "event_type",
            "event_subtype",
            "bpmn_event_type",
        )
    )

    event_definition = normalize_text(
        get_metadata_value(
            "event_definition",
            "event_definition_type",
            "definition_type",
        )
    )

    combined_text = normalize_text(
        " ".join(
            part for part in (
                shape_text,
                shape_name,
                shape_type,
                event_subtype,
                event_definition,
            )
            if part
        )
    )

    # ---------------------------------------------------------
    # 3. NORMALIZE FLOW COUNTS
    # ---------------------------------------------------------

    try:
        incoming_count = max(0, int(incoming_count))
    except (TypeError, ValueError):
        incoming_count = 0

    try:
        outgoing_count = max(0, int(outgoing_count))
    except (TypeError, ValueError):
        outgoing_count = 0

    # ---------------------------------------------------------
    # 4. KEYWORD MATCHING
    # ---------------------------------------------------------

    def contains_keyword(text, keywords):
        """
        Match complete words where possible.
        Persian phrases are matched as normalized substrings.
        """

        if not text:
            return False

        for keyword in keywords:
            keyword = normalize_text(keyword)

            if not keyword:
                continue

            if keyword in text:
                return True

        return False

    # ---------------------------------------------------------
    # 5. EXPLICIT BPMN TYPE METADATA
    # ---------------------------------------------------------

    explicit_type_mapping = {
        "startevent": "startEvent",
        "start event": "startEvent",
        "bpmn:startevent": "startEvent",

        "endevent": "endEvent",
        "end event": "endEvent",
        "bpmn:endevent": "endEvent",

        "intermediatecatchevent": "intermediateCatchEvent",
        "intermediate catch event": "intermediateCatchEvent",

        "intermediatethrowevent": "intermediateThrowEvent",
        "intermediate throw event": "intermediateThrowEvent",

        "boundaryevent": "boundaryEvent",
        "boundary event": "boundaryEvent",
    }

    for candidate in (shape_type, event_subtype):
        if candidate in explicit_type_mapping:
            result = explicit_type_mapping[candidate]

            log_correction(
                f"{metadata.get('name', 'Unnamed Shape')}: "
                f"identified as {result} using explicit metadata."
            )

            return result

    # ---------------------------------------------------------
    # 6. START EVENT KEYWORDS
    # ---------------------------------------------------------

    start_keywords = [
        # Persian
        "شروع",
        "آغاز",
        "ابتدا",
        "شروع فرایند",
        "آغاز فرایند",
        "شروع فرآیند",
        "آغاز فرآیند",
        "نقطه شروع",
        "رویداد شروع",
        "رویداد آغاز",
        "شروع گردش کار",
        "شروع گردش‌کار",
        "آغاز گردش کار",
        "آغاز گردش‌کار",
        "شروع عملیات",
        "شروع فعالیت",
        "شروع کار",
        "ورود به فرایند",
        "ورود به فرآیند",
        "آغاز کار",
        "فعال سازی فرایند",
        "فعال‌سازی فرایند",

        # English
        "start",
        "begin",
        "beginning",
        "start event",
        "process start",
        "process begins",
        "workflow start",
        "workflow begins",
        "initial event",
        "initial state",
        "trigger process",
        "process trigger",
        "entry point",
        "start process",
        "start workflow",
    ]

    # ---------------------------------------------------------
    # 7. END EVENT KEYWORDS
    # ---------------------------------------------------------

    end_keywords = [
        # Persian
        "پایان",
        "اتمام",
        "خاتمه",
        "ختم",
        "تمام",
        "پایان فرایند",
        "پایان فرآیند",
        "اتمام فرایند",
        "اتمام فرآیند",
        "خاتمه فرایند",
        "خاتمه فرآیند",
        "ختم فرایند",
        "ختم فرآیند",
        "پایان گردش کار",
        "پایان گردش‌کار",
        "اتمام گردش کار",
        "اتمام گردش‌کار",
        "پایان عملیات",
        "اتمام عملیات",
        "پایان فعالیت",
        "اتمام فعالیت",
        "پایان کار",
        "تکمیل فرایند",
        "تکمیل فرآیند",
        "فرایند تکمیل شد",
        "فرآیند تکمیل شد",
        "کار تکمیل شد",
        "عملیات تکمیل شد",
        "خروج از فرایند",
        "خروج از فرآیند",
        "نقطه پایان",
        "رویداد پایان",
        "رویداد خاتمه",
        "پایان موفق",
        "تکمیل موفق",
        "تکمیل موفقیت آمیز",
        "تکمیل موفقیت‌آمیز",

        # English
        "end",
        "finish",
        "finished",
        "terminate",
        "terminated",
        "termination",
        "complete",
        "completed",
        "completion",
        "end event",
        "process end",
        "process ends",
        "workflow end",
        "workflow ends",
        "final event",
        "final state",
        "exit point",
        "process complete",
        "process completed",
        "workflow completed",
        "successfully completed",
        "successful completion",
    ]

    # ---------------------------------------------------------
    # 8. INTERMEDIATE THROW EVENT KEYWORDS
    # ---------------------------------------------------------

    throw_keywords = [
        # Persian
        "ارسال پیام",
        "ارسال سیگنال",
        "ارسال اعلان",
        "ارسال هشدار",
        "ارسال درخواست",
        "انتشار پیام",
        "انتشار سیگنال",
        "صدور هشدار",
        "صدور اعلان",
        "ایجاد پیام",
        "تولید پیام",
        "پرتاب سیگنال",
        "پرتاب لینک",
        "ارسال لینک",
        "فعال سازی رویداد",
        "فعال‌سازی رویداد",

        # English
        "throw",
        "throw event",
        "send message",
        "message throw",
        "throw message",
        "send signal",
        "signal throw",
        "throw signal",
        "send notification",
        "send alert",
        "publish message",
        "publish signal",
        "intermediate throw",
        "link throw",
    ]

    # ---------------------------------------------------------
    # 9. INTERMEDIATE CATCH EVENT KEYWORDS
    # ---------------------------------------------------------

    catch_keywords = [
        # Persian
        "دریافت پیام",
        "دریافت سیگنال",
        "دریافت اعلان",
        "دریافت هشدار",
        "انتظار برای پیام",
        "انتظار پیام",
        "انتظار برای پاسخ",
        "انتظار پاسخ",
        "دریافت پاسخ",
        "انتظار برای رویداد",
        "انتظار رویداد",
        "وقوع رویداد",
        "رسیدن پیام",
        "رسیدن پاسخ",
        "شروع زمان سنجی",
        "شروع زمان‌سنجی",
        "انتظار زمانی",
        "توقف تا زمان",
        "توقف تا دریافت",
        "دریافت لینک",

        # English
        "catch",
        "catch event",
        "receive message",
        "message catch",
        "catch message",
        "receive signal",
        "signal catch",
        "catch signal",
        "wait for message",
        "wait for response",
        "wait for signal",
        "wait for event",
        "intermediate catch",
        "timer catch",
        "wait until",
        "link catch",
    ]

    # ---------------------------------------------------------
    # 10. BOUNDARY EVENT KEYWORDS
    # ---------------------------------------------------------

    boundary_keywords = [
        "boundary event",
        "attached event",
        "مرزی",
        "رویداد مرزی",
        "رویداد چسبیده",
        "رویداد متصل به فعالیت",
    ]

    # ---------------------------------------------------------
    # 11. EXPLICIT EVENT DEFINITION KEYWORDS
    # ---------------------------------------------------------

    event_definition_keywords = {
        "timer": [
            "timer",
            "زمان سنجی",
            "زمان‌سنجی",
            "تایمر",
            "زمان",
            "مهلت",
            "ضرب الاجل",
            "ضرب‌الاجل",
            "تاخیر",
            "تأخیر",
            "تاریخ",
            "ساعت",
            "هر روز",
            "هر هفته",
            "هر ماه",
            "پس از مدت",
            "بعد از مدت",
        ],

        "message": [
            "message",
            "پیام",
            "نامه",
            "ایمیل",
            "email",
            "ارسال نامه",
            "دریافت نامه",
            "درخواست",
            "پاسخ",
            "request",
            "response",
        ],

        "signal": [
            "signal",
            "سیگنال",
            "اعلان عمومی",
            "broadcast",
            "انتشار",
        ],

        "error": [
            "error",
            "exception",
            "خطا",
            "استثنا",
            "خطای سیستم",
            "خطای فنی",
            "خطای پردازش",
            "شکست عملیات",
            "عدم موفقیت",
        ],

        "escalation": [
            "escalation",
            "تشدید",
            "ارجاع به مدیر",
            "ارجاع به سطح بالاتر",
            "ارتقا به مدیریت",
        ],

        "cancel": [
            "cancel",
            "cancellation",
            "لغو",
            "کنسل",
            "ابطال",
            "درخواست لغو",
        ],

        "compensation": [
            "compensation",
            "جبران",
            "جبران خسارت",
            "بازپرداخت",
        ],

        "conditional": [
            "condition",
            "conditional",
            "شرط",
            "مشروط",
            "در صورت",
            "به محض اینکه",
            "هرگاه",
            "اگر",
        ],

        "link": [
            "link",
            "لینک",
            "اتصال",
            "پیوند",
        ],

        "multiple": [
            "multiple",
            "چندگانه",
            "چند رویداد",
            "چندین رویداد",
        ],

        "terminate": [
            "terminate",
            "termination",
            "توقف کامل",
            "خاتمه کامل",
            "پایان فوری",
            "توقف کل فرایند",
            "توقف کل فرآیند",
        ],
    }

    # ---------------------------------------------------------
    # 12. EXPLICIT START / END DETECTION
    # ---------------------------------------------------------

    is_start = contains_keyword(combined_text, start_keywords)
    is_end = contains_keyword(combined_text, end_keywords)

    if is_start and not is_end:
        result = "startEvent"

        log_correction(
            f"{metadata.get('name', 'Unnamed Shape')}: "
            "identified as Start Event using text semantics."
        )

        return result

    if is_end and not is_start:
        result = "endEvent"

        log_correction(
            f"{metadata.get('name', 'Unnamed Shape')}: "
            "identified as End Event using text semantics."
        )

        return result

    # ---------------------------------------------------------
    # 13. EXPLICIT BOUNDARY EVENT DETECTION
    # ---------------------------------------------------------

    if contains_keyword(combined_text, boundary_keywords):
        log_correction(
            f"{metadata.get('name', 'Unnamed Shape')}: "
            "identified as Boundary Event using metadata/text."
        )

        return "boundaryEvent"

    # ---------------------------------------------------------
    # 14. THROW / CATCH EVENT DETECTION
    # ---------------------------------------------------------

    is_throw = contains_keyword(combined_text, throw_keywords)
    is_catch = contains_keyword(combined_text, catch_keywords)

    if is_throw and not is_catch:
        log_correction(
            f"{metadata.get('name', 'Unnamed Shape')}: "
            "identified as Intermediate Throw Event using text semantics."
        )

        return "intermediateThrowEvent"

    if is_catch and not is_throw:
        log_correction(
            f"{metadata.get('name', 'Unnamed Shape')}: "
            "identified as Intermediate Catch Event using text semantics."
        )

        return "intermediateCatchEvent"

    # ---------------------------------------------------------
    # 15. GRAPH POSITION DETECTION
    # ---------------------------------------------------------

    if incoming_count == 0 and outgoing_count > 0:
        log_correction(
            f"{metadata.get('name', 'Unnamed Shape')}: "
            "classified as Start Event based on graph position."
        )

        return "startEvent"

    if outgoing_count == 0 and incoming_count > 0:
        log_correction(
            f"{metadata.get('name', 'Unnamed Shape')}: "
            "classified as End Event based on graph position."
        )

        return "endEvent"

    # ---------------------------------------------------------
    # 16. AMBIGUOUS / ISOLATED EVENT DETECTION
    # ---------------------------------------------------------

    if incoming_count == 0 and outgoing_count == 0:

        # A named isolated event cannot reliably be identified
        # as Start or End from graph position alone.

        log_correction(
            f"{metadata.get('name', 'Unnamed Shape')}: "
            "isolated event classified as Start Event by fallback."
        )

        return "startEvent"

    # ---------------------------------------------------------
    # 17. DEFAULT INTERMEDIATE EVENT
    # ---------------------------------------------------------

    if outgoing_count > 0:
        log_correction(
            f"{metadata.get('name', 'Unnamed Shape')}: "
            "classified as Intermediate Throw Event by graph position."
        )

        return "intermediateThrowEvent"

    log_correction(
        f"{metadata.get('name', 'Unnamed Shape')}: "
        "classified as Intermediate Catch Event by fallback."
    )

    return "intermediateCatchEvent"


def determine_gateway_type(metadata, incoming_count, outgoing_count):
    """
    Without explicit BPMN markers, a diamond is conservatively
    represented as an Exclusive Gateway.

    This is a default mapping, not proof of the intended semantics.
    """
    log_correction(
        f"{metadata['name']}: defaulted to Exclusive Gateway."
    )

    return "exclusiveGateway"


def determine_gateway_direction(incoming_count, outgoing_count):
    if incoming_count > 1 and outgoing_count > 1:
        return "Mixed"

    if incoming_count > 1:
        return "Converging"

    if outgoing_count > 1:
        return "Diverging"

    return "Unspecified"


def add_gateway_direction(gateway_element, incoming_count, outgoing_count):
    gateway_element.set(
        "gatewayDirection",
        determine_gateway_direction(
            incoming_count,
            outgoing_count,
        ),
    )


# ============================================================
# MAIN PAGE CONVERSION
# ============================================================

def convert_page(page, root, page_number=1):
    page_name = str(page.Name)

    page_width, page_height = get_page_size(page)

    all_shapes = get_page_shapes(page)

    metadata_list = [
        get_shape_metadata(shape)
        for shape in all_shapes
    ]

    nodes = []
    connectors = []

    for metadata in metadata_list:

        kind = classify_shape(metadata)

        metadata["kind"] = kind

        if kind == "connector":
            connectors.append(metadata)

        else:
            nodes.append(metadata)

    # Ignore empty shapes without meaningful labels.
    # They may still be useful as containers, but this initial
    # version focuses on process nodes and connectors.
    nodes = [
        node for node in nodes
        if node["text"]
        or node["kind"] in {"event", "gateway"}
    ]

    if not nodes:
        log_warning(
            f"Page '{page_name}' contains no identifiable process nodes."
        )
        return

    process_id = safe_id(
        f"Process_Page_{page_number}",
        "Process",
    )

    process = add_process(
        root,
        process_id,
        page_name,
    )

    plane = create_diagram(
        root,
        process_id,
    )

    # Create a stable mapping between Visio shape IDs and BPMN IDs.
    shape_to_node = {}
    node_metadata = {}

    for index, node in enumerate(nodes, start=1):

        shape = node["shape"]

        try:
            visio_id = str(shape.ID)
        except Exception:
            visio_id = str(index)

        bpmn_id = safe_id(
            f"Node_{page_number}_{visio_id}",
            "Node",
        )

        node["bpmn_id"] = bpmn_id
        node["visio_id"] = visio_id

        shape_to_node[visio_id] = node
        node_metadata[bpmn_id] = node

    # Resolve connector endpoints.
    flow_records = []

    for index, connector in enumerate(connectors, start=1):

        shape = connector["shape"]

        begin_shape, end_shape = get_glued_endpoints(shape)

        start_point, end_point = get_connector_endpoints(shape)

        reverse_direction = get_connector_direction(shape)

        if reverse_direction:
            begin_shape, end_shape = end_shape, begin_shape
            start_point, end_point = end_point, start_point

        source_node = None
        target_node = None

        if begin_shape is not None:
            try:
                source_id = str(begin_shape.ID)

                source_node = shape_to_node.get(source_id)

            except Exception:
                pass

        if end_shape is not None:
            try:
                target_id = str(end_shape.ID)

                target_node = shape_to_node.get(target_id)

            except Exception:
                pass

        # Fallback for connectors that are not glued.
        if source_node is None:
            source_node = find_nearest_node(
                start_point,
                nodes,
            )

        if target_node is None:
            target_node = find_nearest_node(
                end_point,
                nodes,
            )

        if source_node is None or target_node is None:
            log_warning(
                f"Connector '{connector['name']}' "
                f"could not be connected to both endpoints."
            )
            continue

        if source_node["bpmn_id"] == target_node["bpmn_id"]:
            log_warning(
                f"Connector '{connector['name']}' "
                f"appears to connect a node to itself."
            )

        flow_records.append({
            "metadata": connector,
            "source": source_node,
            "target": target_node,
            "start": start_point,
            "end": end_point,
            "id": safe_id(
                f"Flow_{page_number}_{index}",
                "Flow",
            ),
            "text": get_connector_text(shape),
            "dashed": is_dashed_connector(shape),
        })

    incoming = defaultdict(int)
    outgoing = defaultdict(int)

    for flow in flow_records:
        if flow["dashed"]:
            continue

        incoming[flow["target"]["bpmn_id"]] += 1
        outgoing[flow["source"]["bpmn_id"]] += 1

    # Create BPMN semantic nodes.
    for node in nodes:

        node_id = node["bpmn_id"]
        kind = node["kind"]
        name = node["text"]

        if kind == "event":

            event_type = determine_event_type(
                node,
                incoming[node_id],
                outgoing[node_id],
            )

            add_bpmn_node(
                process,
                event_type,
                node_id,
                name,
            )

        elif kind == "gateway":

            gateway_type = determine_gateway_type(
                node,
                incoming[node_id],
                outgoing[node_id],
            )

            gateway = add_bpmn_node(
                process,
                gateway_type,
                node_id,
                name,
            )

            add_gateway_direction(
                gateway,
                incoming[node_id],
                outgoing[node_id],
            )

        elif kind == "annotation":

            log_warning(
                f"Annotation '{name}' is not yet converted "
                "as a dedicated BPMN TextAnnotation."
            )

            # Conservative fallback: represent it as a task
            # rather than silently discarding its text.
            add_bpmn_node(
                process,
                "task",
                node_id,
                name,
            )

        else:

            add_bpmn_node(
                process,
                "task",
                node_id,
                name,
            )

        add_di_shape(
            plane,
            node_id,
            node,
            page_height,
        )

        REPORT["elements"] += 1

    # Create BPMN flows.
    for flow in flow_records:

        if flow["dashed"]:

            # Association is legal within a process.
            # Message Flow requires separate participants.
            log_warning(
                f"Dashed connector '{flow['metadata']['name']}' "
                "was mapped to Association. "
                "Message Flow requires participant information."
            )

            continue

        flow_id = flow["id"]

        source_id = flow["source"]["bpmn_id"]
        target_id = flow["target"]["bpmn_id"]

        sequence = add_sequence_flow(
            process,
            flow_id,
            source_id,
            target_id,
            flow["text"],
        )

        # Connector labels leaving gateways are interpreted
        # as outgoing condition labels.
        if flow["source"]["kind"] == "gateway" and flow["text"]:

            condition = ET.SubElement(
                sequence,
                qname(BPMN_NS, "conditionExpression"),
                {
                    "xsi:type": "bpmn:tFormalExpression",
                },
            )

            condition.text = flow["text"]

        add_di_edge(
            plane,
            flow_id,
            flow["source"],
            flow["target"],
        )

        REPORT["flows"] += 1

    # Validate gateway default behavior.
    for node in nodes:

        if node["kind"] != "gateway":
            continue

        node_id = node["bpmn_id"]

        if outgoing[node_id] > 1:

            outgoing_flows = [
                flow
                for flow in flow_records
                if (
                    not flow["dashed"]
                    and flow["source"]["bpmn_id"] == node_id
                )
            ]

            unlabeled = [
                flow
                for flow in outgoing_flows
                if not flow["text"]
            ]

            if unlabeled:
                log_warning(
                    f"Gateway '{node['text']}' has outgoing "
                    "paths without labels or conditions."
                )

    # Detect isolated nodes.
    for node in nodes:

        node_id = node["bpmn_id"]

        if (
            incoming[node_id] == 0
            and outgoing[node_id] == 0
        ):
            log_warning(
                f"Isolated BPMN node: {node['text']}"
            )

    # Detect a process with no Start Event.
    starts = [
        node
        for node in nodes
        if (
            node["kind"] == "event"
            and incoming[node["bpmn_id"]] == 0
        )
    ]

    if not starts:
        log_warning(
            f"Page '{page_name}' has no identifiable Start Event."
        )

    # Detect a process with no End Event.
    ends = [
        node
        for node in nodes
        if (
            node["kind"] == "event"
            and outgoing[node["bpmn_id"]] == 0
        )
    ]

    if not ends:
        log_warning(
            f"Page '{page_name}' has no identifiable End Event."
        )


# ============================================================
# BPMN XML VALIDATION
# ============================================================

def validate_generated_xml(root):
    """
    Perform basic structural checks on the generated BPMN XML.
    """
    errors = []

    process_elements = root.findall(
        "bpmn:process",
        NS,
    )

    if not process_elements:
        errors.append(
            "No BPMN process was generated."
        )

    for process in process_elements:

        node_ids = set()

        for child in list(process):

            child_type = child.tag.split("}")[-1]

            if child_type in {
                "startEvent",
                "endEvent",
                "intermediateCatchEvent",
                "intermediateThrowEvent",
                "task",
                "userTask",
                "serviceTask",
                "manualTask",
                "scriptTask",
                "businessRuleTask",
                "sendTask",
                "receiveTask",
                "exclusiveGateway",
                "parallelGateway",
                "inclusiveGateway",
                "eventBasedGateway",
                "complexGateway",
            }:

                node_ids.add(child.get("id"))

        for flow in process.findall(
            "bpmn:sequenceFlow",
            NS,
        ):

            source = flow.get("sourceRef")
            target = flow.get("targetRef")

            if source not in node_ids:
                errors.append(
                    f"SequenceFlow {flow.get('id')} "
                    f"has an invalid sourceRef: {source}"
                )

            if target not in node_ids:
                errors.append(
                    f"SequenceFlow {flow.get('id')} "
                    f"has an invalid targetRef: {target}"
                )

    return errors


# ============================================================
# REPORT
# ============================================================

def save_report(output_file):
    report_path = Path(output_file).with_suffix(
        ".report.txt"
    )

    lines = [
        APP_NAME,
        f"Version: {VERSION}",
        "",
        "CONVERSION SUMMARY",
        "------------------",
        f"Converted elements: {REPORT['elements']}",
        f"Converted flows: {REPORT['flows']}",
        f"Warnings: {len(REPORT['warnings'])}",
        f"Automatic corrections: {len(REPORT['corrections'])}",
        f"Errors: {len(REPORT['errors'])}",
        "",
        "AUTOMATIC CORRECTIONS",
        "---------------------",
    ]

    lines.extend(
        REPORT["corrections"] or ["None"]
    )

    lines.extend([
        "",
        "WARNINGS",
        "--------",
    ])

    lines.extend(
        REPORT["warnings"] or ["None"]
    )

    lines.extend([
        "",
        "ERRORS",
        "------",
    ])

    lines.extend(
        REPORT["errors"] or ["None"]
    )

    report_path.write_text(
        "\n".join(lines),
        encoding="utf-8",
    )

    return report_path


# ============================================================
# MAIN CONVERSION
# ============================================================

def convert_visio_to_bpmn(source_file, output_file=None):
    source_file = Path(source_file)

    if not source_file.exists():
        raise FileNotFoundError(
            f"Visio file not found: {source_file}"
        )

    if output_file is None:
        output_file = source_file.with_suffix(".bpmn")

    output_file = Path(output_file)

    output_file.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    # Reset report.
    for key in REPORT:
        if isinstance(REPORT[key], list):
            REPORT[key].clear()
        else:
            REPORT[key] = 0

    app = None
    document = None

    try:
        print("Starting Microsoft Visio...")

        app = win32com.client.DispatchEx(
            "Visio.Application"
        )

        app.Visible = False

        print("Opening Visio document...")

        document = app.Documents.Open(
            str(source_file.resolve())
        )

        root = create_definitions()

        page_count = document.Pages.Count

        if page_count == 0:
            raise ValueError(
                "The Visio document contains no pages."
            )

        for index in range(1, page_count + 1):

            page = document.Pages.Item(index)

            print(
                f"Analyzing page {index}/{page_count}: "
                f"{page.Name}"
            )

            try:
                convert_page(
                    page,
                    root,
                    page_number=index,
                )

            except Exception as exc:
                message = (
                    f"Page '{page.Name}' conversion failed: {exc}"
                )

                REPORT["errors"].append(message)

                print(message)

        validation_errors = validate_generated_xml(root)

        REPORT["errors"].extend(validation_errors)

        # Write XML even when warnings exist.
        # Do not write a misleading file if structural validation fails.
        if validation_errors:
            raise ValueError(
                "Generated BPMN XML failed structural validation. "
                "See the conversion report."
            )

        tree = ET.ElementTree(root)

        try:
            ET.indent(
                tree,
                space="    ",
            )
        except AttributeError:
            pass

        tree.write(
            str(output_file),
            encoding="utf-8",
            xml_declaration=True,
        )

        report_path = save_report(output_file)

        print("\nCONVERSION COMPLETED")
        print(f"BPMN file: {output_file}")
        print(f"Report: {report_path}")

        return output_file, report_path

    except Exception as exc:

        REPORT["errors"].append(
            f"{exc}\n{traceback.format_exc()}"
        )

        save_report(output_file)

        raise

    finally:

        # Close only the document opened by this converter.
        if document is not None:
            try:
                document.Close()
            except Exception:
                pass

        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass


# ============================================================
# FILE SELECTION
# ============================================================

def choose_input_file():
    import tkinter as tk
    from tkinter import filedialog

    window = None

    try:
        window = tk.Tk()
        window.withdraw()
        window.update_idletasks()
        window.attributes("-topmost", True)

        filename = filedialog.askopenfilename(
            parent=window,
            title="Select Visio Drawing",
            filetypes=[
                ("Visio Files", "*.vsdx;*.vsd"),
                ("Visio 2003-2010", "*.vsd"),
                ("All Files", "*.*"),
            ],
        )

        return filename or ""

    except Exception as exc:
        print(f"File selection error: {exc}")
        return ""

    finally:
        if window is not None:
            try:
                window.destroy()
            except Exception:
                pass


def choose_output_file(default_path):
    import tkinter as tk
    from tkinter import filedialog

    window = None

    try:
        window = tk.Tk()
        window.withdraw()
        window.update_idletasks()
        window.attributes("-topmost", True)

        filename = filedialog.asksaveasfilename(
            parent=window,
            title="Save BPMN 2.0 File",
            initialfile=Path(default_path).name,
            defaultextension=".bpmn",
            filetypes=[
                ("BPMN 2.0 Files", "*.bpmn"),
                ("XML Files", "*.xml"),
            ],
        )

        return filename or ""

    except Exception as exc:
        print(f"Output selection error: {exc}")
        return ""

    finally:
        if window is not None:
            try:
                window.destroy()
            except Exception:
                pass


# ============================================================
# USER INTERFACE
# ============================================================

def main():
    print("\n" + "=" * 65)
    print(APP_NAME)
    print(f"Version {VERSION}")
    print("=" * 65)

    print("\n1. Convert Visio to BPMN 2.0")
    print("2. Exit")

    choice = input("\nSelect option: ").strip()

    if choice == "2":
        return

    if choice != "1":
        print("Invalid option.")
        return

    source = choose_input_file()

    if not source:
        source = input(
            "\nEnter Visio file path: "
        ).strip().strip('"')

    if not source:
        print("No input file selected.")
        return

    default_output = str(
        Path(source).with_suffix(".bpmn")
    )

    output = choose_output_file(
        default_output
    )

    if not output:
        output = input(
            "\nEnter output BPMN path "
            "(press ENTER for default): "
        ).strip().strip('"')

    if not output:
        output = default_output

    try:
        bpmn_file, report_file = convert_visio_to_bpmn(
            source,
            output,
        )

        print("\n" + "=" * 65)
        print("CONVERSION REPORT")
        print("=" * 65)

        print(
            f"Elements: {REPORT['elements']}"
        )

        print(
            f"Flows: {REPORT['flows']}"
        )

        print(
            f"Corrections: {len(REPORT['corrections'])}"
        )

        print(
            f"Warnings: {len(REPORT['warnings'])}"
        )

        print(
            f"Errors: {len(REPORT['errors'])}"
        )

        print("\nBPMN:", bpmn_file)
        print("Report:", report_file)

    except Exception as exc:

        print("\nCONVERSION FAILED")
        print(str(exc))

        traceback.print_exc()

    input("\nPress ENTER to exit...")


if __name__ == "__main__":
    main()
