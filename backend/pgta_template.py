"""
PGT-A Report Template Engine
Generates PDF and DOCX reports from patient and embryo data
"""

from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch, mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, 
    Spacer, PageBreak, Image, KeepTogether, CondPageBreak
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY, TA_RIGHT
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from PIL import Image as PILImage
import os
import sys
import base64
from io import BytesIO
from datetime import datetime
from pgta_assets import HEADER_LOGO_B64, FOOTER_BANNER_B64, SIGN_ANAND_B64, SIGN_SACHIN_B64, SIGN_DIRECTOR_B64


class PGTAReportTemplate:
    """Template engine for PGT-A reports"""
    
    COLORS = {
        'patient_info_bg': '#F1F1F7',
        'results_header_bg': '#F9BE8F',
        'grey_bg': '#F2F2F2',
        'blue_title': '#1F497D',
        'approval_blue': '#4F81BD'
    }
    
    # Page dimensions (US Letter)
    PAGE_WIDTH = 612  # points
    PAGE_HEIGHT = 792  # points
    MARGIN_LEFT = 58  # Reduced from 72 to allow title in single line
    MARGIN_RIGHT = 58
    MARGIN_TOP = 70
    MARGIN_BOTTOM = 60
    CONTENT_WIDTH = PAGE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT  # 496 points
    
    # Asset paths
    ASSETS_DIR = "assets/pgta"
    HEADER_LOGO = os.path.join(ASSETS_DIR, "image_page1_0.png")
    FOOTER_BANNER = os.path.join(ASSETS_DIR, "image_page1_1.png")
    FOOTER_LOGO = os.path.join(ASSETS_DIR, "image_page1_2.png")
    
    # Static content
    METHODOLOGY_TEXT = """Chromosomal aneuploidy analysis was performed using ChromInst® PGT-A kit from Yikon Genomics (Suzhou) Co., Ltd - China. The Yikon - ChromInst® PGT-A kit with the Genemind - SURFSeq 5000* High-throughput Sequencing Platform allows detection of aneuploidies in all 23 sets of Chromosomes. Probes are not covering the p arm of acrocentric chromosomes as they are rich in repeat regions and RNA markers and devoid of genes. Changes in this region will not be detected. However, these regions have less clinical significance due to the absence of genes. Chromosomal aneuploidy can be detected by copy number variations (CNVs), which represent a class of variation in which segments of the genome have been duplicated (gains) or deleted (losses). Large, genomic copy number imbalances can range from sub-chromosomal regions to entire chromosomes. Inherited and de-novo CNVs (up to 10 Mb) have been associated with many disease conditions. This assay was performed on DNA extracted from embryo biopsy samples."""
    
    MOSAICISM_TEXT = """Mosaicism arises in the embryo due to mitotic errors which lead to the production of karyotypically distinct cell lineages within a single embryo [1]. NGS has the sensitivity to detect mosaicism when 30% or the above cells are abnormal [2]. Mosaicism is reported in our laboratory as follows [3]."""
    
    MOSAICISM_BULLETS = [
        "Embryos with less than 30% mosaicism are considered as euploid.",
        "Embryos with 30% to 50% mosaicism will be reported as low level mosaic, 51% to 80% mosaicism will be reported as high level mosaic.",
        "When three chromosomes or more than three chromosomes showing mosaic change, it will be denoted as complex mosaic.",
        "If greater than 80% mosaicism detected in an embryo it will be considered aneuploid."
    ]
    
    MOSAICISM_CLINICAL = """Clinical significance of transferring mosaic embryos is still under evaluation. Based on Preimplantation Genetic Diagnosis International Society (PGDIS) Position Statement – 2019 transfer of these embryos should be considered only after appropriate counselling of the patient and alternatives have been discussed. Invasive prenatal testing with karyotyping in the amniotic fluid needs to be advised in such cases [4]. As shown in published literature evidence, such transfers can result in normal pregnancy or miscarriage or an offspring with chromosomal mosaicism [5,6,7]."""
    
    LIMITATIONS = [
        "This technique cannot detect point mutations, balanced translocations, inversions, triploidy, uniparental disomy and epigenetic modifications.",
        "Probes used do not cover the p arm of acrocentric chromosomes as they are rich in repeat regions and RNA markers and devoid of genes. Changes in this region will not be detected. However, these regions have less clinical significance due to the absence of genes.",
        "Deletions and duplications with the size of < 10 Mb cannot be detected.",
        "Risk of misinterpretation of the actual embryo karyotype due to the presence of chromosomal mosaicism, either at cleavage-stage or at blastocyst stage may exist.",
        "This technique cannot detect variants of polyploidy and haploidy",
        "NGS without genotyping cannot identify the nature (meiotic or mitotic) nor the parental origin of aneuploidies",
        "Due to the intrinsic nature of chromosomal mosaicism, the chromosomal make-up achieved from a biopsy only may represent a picture of a small part of the embryo and may not necessarily reflect the chromosomal content of the entire embryo. Also, the mosaicism level inferred from a multi-cell TE biopsy might not unequivocally represent the exact chromosomal mosaicism percentage of the TE cells or the inner cell mass constitution."
    ]
    
    REFERENCES = [
        'McCoy, Rajiv C. "Mosaicism in Preimplantation human embryos: when chromosomal abnormalities are the norm." Trends in genetics 33.7 (2017): 448-463.',
        'ESHRE PGT-SR/PGT-A Working Group, et al. "ESHRE PGT Consortium good practice recommendations for the detection of structural and numerical chromosomal aberrations." Human reproduction open 2020.3 (2020): hoaa017.',
        'ESHRE Working Group on Chromosomal Mosaicism, et al. "ESHRE survey results and good practice recommendations on managing chromosomal mosaicism." Hum Reprod Open. 2022 Nov 7;2022(4):hoac044.',
        'Cram, D. S., et al. "PGDIS position statement on the transfer of mosaic embryos 2019." Reproductive biomedicine online 39 (2019): e1-e4.',
        'Victor, Andrea R., et al. "One hundred mosaic embryos transferred prospectively in a single clinic: exploring when and why they result in healthy pregnancies." Fertility and sterility 111.2 (2019): 280-293.',
        'Lin, Pin-Yao, et al. "Clinical outcomes of single mosaic embryo transfer: high-level or low-level mosaic embryo, does it matter?" Journal of clinical medicine 9.6 (2020): 1695.',
        'Kahraman, Semra, et al. "The birth of a baby with mosaicism resulting from a known mosaic embryo transfer: a case report." Human Reproduction 35.3 (2020): 727-733.'
    ]
    
    SIGNATURES = [
        {"name": "Anand Babu. K, Ph.D", "title": "Molecular Biologist"},
        {"name": "Sachin D Honguntikar, Ph.D", "title": "Molecular Geneticist"},
        {"name": "Dr Suriyakumar G", "title": "Director"}
    ]
    
    @staticmethod
    def get_resource_path(relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        if hasattr(sys, '_MEIPASS'):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
            
        # Join and then use abspath to normalize the path (fix slash direction etc)
        return os.path.abspath(os.path.join(base_path, relative_path))

    def __init__(self, assets_dir="assets/pgta"):
        """Initialize template with asset directory"""
        # Resolve the assets directory relative to the script location
        self.ASSETS_DIR = self.get_resource_path(assets_dir)
        print(f"INFO: Assets Directory resolved to: {self.ASSETS_DIR}")
        
        # Hardcode specific asset filenames
        self.HEADER_LOGO = os.path.join(self.ASSETS_DIR, "image_page1_0.png")
        self.FOOTER_BANNER = os.path.join(self.ASSETS_DIR, "image_page1_1.png")
        self.FOOTER_LOGO = os.path.join(self.ASSETS_DIR, "image_page1_2.png")
        self.GENQA_LOGO = os.path.join(self.ASSETS_DIR, "genqa_logo.png")
        self.SIGNS_IMAGE = os.path.join(self.ASSETS_DIR, "signs.png")
        
        # Verify critical assets
        for label, path in [("Header", self.HEADER_LOGO), ("Footer", self.FOOTER_BANNER), ("Signs", self.SIGNS_IMAGE)]:
            if not os.path.exists(path):
                print(f"CRITICAL: {label} missing at {path}")
            else:
                print(f"FOUND: {label} ({os.path.getsize(path)} bytes)")
        
        # Create custom styles
        self.styles = getSampleStyleSheet()
        self._register_fonts()
        self._create_custom_styles()
    
    def _register_fonts(self):
        """Register custom fonts if they exist in assets/fonts"""
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.pdfbase.pdfmetrics import registerFontFamily
        
        fonts_dir = os.path.join(self.ASSETS_DIR, "fonts")
        if not os.path.exists(fonts_dir):
            return
            
        # Actual filenames from user upload
        font_configs = [
            # Segoe UI
            {'name': 'SegoeUI', 'file': 'SEGOEUI.TTF'},
            {'name': 'SegoeUI-Bold', 'file': 'SEGOEUIB.TTF'},
            {'name': 'SegoeUI-Italic', 'file': 'SEGOEUII.TTF'},
            {'name': 'SegoeUI-BoldItalic', 'file': 'SEGOEUIZ.TTF'},
            {'name': 'SegoeUI-Semibold', 'file': 'SEGUISB.TTF'},
            {'name': 'SegoeUI-SemiboldItalic', 'file': 'SEGUISBI.TTF'},
            # Gill Sans MT
            {'name': 'GillSansMT', 'file': 'GIL_____.TTF'},
            {'name': 'GillSansMT-Bold', 'file': 'GILB____.TTF'},
            {'name': 'GillSansMT-Italic', 'file': 'GILI____.TTF'},
            {'name': 'GillSansMT-BoldItalic', 'file': 'GILBI___.TTF'},
            # Calibri
            {'name': 'Calibri', 'file': 'CALIBRI.TTF'},
            {'name': 'Calibri-Bold', 'file': 'CALIBRIB.TTF'},
            {'name': 'Calibri-Italic', 'file': 'CALIBRII.TTF'},
            {'name': 'Calibri-BoldItalic', 'file': 'CALIBRIZ.TTF'},
        ]
        
        registered = []
        for cfg in font_configs:
            font_path = os.path.join(fonts_dir, cfg['file'])
            if os.path.exists(font_path):
                try:
                    pdfmetrics.registerFont(TTFont(cfg['name'], font_path))
                    registered.append(cfg['name'])
                    print(f"Registered font: {cfg['name']}")
                except Exception as e:
                    print(f"Error registering font {cfg['name']}: {e}")
        
        # Register font families if components exist
        if 'SegoeUI' in registered and 'SegoeUI-Bold' in registered:
            registerFontFamily('SegoeUI', normal='SegoeUI', bold='SegoeUI-Bold', italic='SegoeUI-Italic', boldItalic='SegoeUI-BoldItalic')
        if 'GillSansMT' in registered and 'GillSansMT-Bold' in registered:
            registerFontFamily('GillSansMT', normal='GillSansMT', bold='GillSansMT-Bold', italic='GillSansMT-Italic', boldItalic='GillSansMT-BoldItalic')
        if 'Calibri' in registered and 'Calibri-Bold' in registered:
            registerFontFamily('Calibri', normal='Calibri', bold='Calibri-Bold', italic='Calibri-Italic', boldItalic='Calibri-BoldItalic')
    
    def _get_font(self, name, fallback):
        """Helper to get best available font"""
        try:
            from reportlab.pdfbase import pdfmetrics
            pdfmetrics.getFont(name)
            return name
        except:
            return fallback

    def _create_custom_styles(self):
        """Create custom paragraph styles"""

        # Title style
        self.styles.add(ParagraphStyle(
            name='PGTAReportTitle',
            parent=self.styles['Heading1'],
            fontSize=16, # Adjusted to match source and ensure single line
            leading=18,
            textColor=colors.HexColor(self.COLORS['blue_title']),
            alignment=TA_CENTER,
            spaceAfter=12,
            fontName=self._get_font('GillSansMT-Bold', 'Helvetica-Bold')
        ))
        
        # Section header
        self.styles.add(ParagraphStyle(
            name='PGTASectionHeader',
            parent=self.styles['Heading2'],
            fontSize=11,
            leading=13,
            textColor=colors.HexColor(self.COLORS['blue_title']), # Color as in source
            spaceBefore=12,
            spaceAfter=3, # Reduced to accommodate line
            keepWithNext=True,
            fontName=self._get_font('SegoeUI-Bold', 'Helvetica-Bold')
        ))
        
        # Body text
        self.styles.add(ParagraphStyle(
            name='PGTABodyText',
            parent=self.styles['Normal'],
            fontSize=11,  # Matches source 11.04pt
            leading=13,
            alignment=TA_JUSTIFY,
            fontName=self._get_font('Calibri', 'Helvetica')
        ))
        
        # Small text
        self.styles.add(ParagraphStyle(
            name='PGTASmallText',
            parent=self.styles['Normal'],
            fontSize=8,
            leading=10,
            fontName=self._get_font('SegoeUI', 'Helvetica')
        ))
        
        # Bold disclaimer (PNDT) - Light Bold (Segoe UI Semibold) as requested
        self.styles.add(ParagraphStyle(
            name='PGTADisclaimer',
            parent=self.styles['Normal'],
            fontSize=10.5, 
            leading=12,
            alignment=TA_CENTER,
            fontName=self._get_font('SegoeUI-SemiboldItalic', 'Helvetica-BoldOblique'),
            textColor=colors.black
        ))
        
        # Bullet style
        self.styles.add(ParagraphStyle(
            name='PGTABulletText',
            parent=self.styles['Normal'],
            fontSize=11, # Increased to match Methodology (11pt)
            leading=13, # Increased leading
            leftIndent=20,
            bulletIndent=10,
            alignment=TA_JUSTIFY,
            fontName=self._get_font('Calibri', 'Helvetica')
        ))
        
        # Signature Approval line style
        self.styles.add(ParagraphStyle(
            name='PGTASigApproval',
            parent=self.styles['Normal'],
            fontSize=12.48, # Exact source size
            leading=14.5,
            textColor=colors.HexColor(self.COLORS['approval_blue']),
            fontName=self._get_font('SegoeUI-Bold', 'Helvetica-Bold')
        ))
        
        # Centered Body style for reliable table alignment
        self.styles.add(ParagraphStyle(
            name='PGTACenteredBodyText',
            parent=self.styles['PGTABodyText'],
            alignment=TA_CENTER
        ))
        
        # Left-aligned Body style for patient info values (no justify gaps)
        self.styles.add(ParagraphStyle(
            name='PGTALeftBodyText',
            parent=self.styles['PGTABodyText'],
            alignment=TA_LEFT
        ))
        
        # Label text style (Force RIGHT alignment, NO justification)
        self.styles.add(ParagraphStyle(
            name='PGTALabelText',
            parent=self.styles['Normal'],
            fontSize=10, 
            leading=12,
            alignment=TA_LEFT,
            wordWrap='CJK',
            fontName=self._get_font('SegoeUI-Bold', 'Helvetica-Bold')
        ))

        # Label text style (Force RIGHT alignment)
        self.styles.add(ParagraphStyle(
            name='PGTALabelTextRight',
            parent=self.styles['Normal'],
            fontSize=10, 
            leading=12,
            alignment=TA_RIGHT,
            wordWrap='CJK',
            fontName=self._get_font('SegoeUI-Bold', 'Helvetica-Bold')
        ))

        # Banner Value style (Matches Label metrics for alignment)
        self.styles.add(ParagraphStyle(
            name='PGTABannerValueText',
            parent=self.styles['Normal'],
            fontSize=10, 
            leading=12,
            alignment=TA_LEFT,
            fontName=self._get_font('SegoeUI-Bold', 'Helvetica-Bold')
        ))
    
    def _get_grid_style(self):
        """Return grid style if enabled"""
        if hasattr(self, 'show_grid') and self.show_grid:
            return [('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#E0E0E0"))]
        return []

    def generate_pdf(self, output_path, patient_data, embryos_data, show_logo=True, show_grid=False):
        """Generate PGT-A report PDF"""
        self.show_grid = show_grid
        # Create PDF document
        doc = SimpleDocTemplate(
            output_path,
            pagesize=letter,
            leftMargin=self.MARGIN_LEFT,
            rightMargin=self.MARGIN_RIGHT,
            topMargin=self.MARGIN_TOP,
            bottomMargin=self.MARGIN_BOTTOM
        )
        
        # Store show_logo preference for the canvas callback
        self._show_logo = show_logo
        
        # Build story (content)
        story = []
        
        # Page 1: Cover with patient info and results summary
        story.extend(self._build_cover_page(patient_data, embryos_data))
        # No PageBreak here to allow Methodology to flow immediately after Results
        
        # Methodolgy content (Flows after cover, includes CondPageBreak for References)
        story.extend(self._build_methodology_page())
        
        # Check if ALL embryos are "Low DNA concentration"
        # If all are Low DNA, we skip all individual pages and add signature to Page 2
        all_low_dna = True if embryos_data else False
        for embryo in embryos_data:
            interp = str(embryo.get('interpretation', '')).upper()
            res = str(embryo.get('result_summary', '')).upper()
            if "LOW DNA" not in interp and "LOW DNA" not in res:
                all_low_dna = False
                break

        if all_low_dna:
            # For all Low DNA embryos, append signature directly after Methodology/References
            story.append(Spacer(1, 24))
            story.append(self._create_signature_table())
        else:
            # Strict page break before embryo results as requested
            story.append(PageBreak())
        
        # Pages 3+: Individual embryo results
        for idx, embryo in enumerate(embryos_data):
            # Skip if Low DNA
            interp = str(embryo.get('interpretation', '')).upper()
            res = str(embryo.get('result_summary', '')).upper()
            if "LOW DNA" in interp or "LOW DNA" in res:
                continue
                
            story.extend(self._build_embryo_page(patient_data, embryo))
            # Add signature under EACH embryo detail section as requested
            story.append(Spacer(1, 12))
            story.append(self._create_signature_table())
            story.append(PageBreak())
        
        # Build PDF
        doc.build(story, onFirstPage=self._add_header_footer, 
                  onLaterPages=self._add_header_footer)
        
        return output_path
    
    def _add_header_footer(self, canvas, doc):
        """Add header and footer to each page using Base64 assets for robustness"""
        show_logo = getattr(self, '_show_logo', True)
            
        canvas.saveState()
        
        def draw_b64_img(b64_str, x, y, w, h):
            try:
                img_data = base64.b64decode(b64_str)
                img = PILImage.open(BytesIO(img_data))
                canvas.drawInlineImage(img, x, y, width=w, height=h, preserveAspectRatio=True)
                return True
            except Exception as e:
                print(f"Error drawing Base64 image: {e}")
                return False

        if show_logo:
            # Draw Header Logo (Base64)
            draw_b64_img(HEADER_LOGO_B64, 72, 720, 468, 72)
            
            # Draw Footer Banner (Base64)
            draw_b64_img(FOOTER_BANNER_B64, 72, 0.4, 468, 66)
        
        # ALWAYS Draw GenQA Logo (per user request: "without logo genqa must be there")
        if os.path.exists(self.GENQA_LOGO):
             try:
                 canvas.drawImage(self.GENQA_LOGO, 454, 35, width=67, height=36, preserveAspectRatio=True, mask='auto')
             except:
                 pass
        
        canvas.restoreState()
    
    def _build_cover_page(self, patient_data, embryos_data):
        """Build cover page with patient info and results summary"""
        elements = []
        
        # Title - Blue color as in source PDF
        title_style = self.styles['PGTAReportTitle']
        title = Paragraph(
            "Preimplantation Genetic Testing for Aneuploidies (PGT-A)",
            title_style
        )
        elements.append(title)
        elements.append(Spacer(1, 6)) # Fixed reduced spacer
        
        # Patient information table
        patient_table = self._create_patient_info_table(patient_data)
        elements.append(patient_table)
        elements.append(Spacer(1, 12))
        
        # PNDT Disclaimer in a grey box
        disclaimer = Paragraph(
            "<b>This test does not reveal sex of the fetus & confers to PNDT act, 1994</b>",
            self.styles['PGTADisclaimer']
        )
        # Use a single-cell table for the background color (Clean white with line as requested)
        disclaimer_table = Table([[disclaimer]], colWidths=[490], hAlign='CENTER')
        disclaimer_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(self.COLORS['grey_bg'])),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ] + self._get_grid_style()))
        elements.append(KeepTogether(disclaimer_table))
        elements.append(Spacer(1, 12))
        
        # Indication
        if 'indication' in patient_data and patient_data['indication']:
            elements.append(self._create_section_header("Indication"))
            elements.append(Spacer(1, 8)) # Gap after header line
            indication_text = Paragraph(patient_data['indication'], self.styles['PGTABodyText'])
            elements.append(indication_text)
            elements.append(Spacer(1, 12))
        
        # Results summary header
        elements.append(self._create_section_header("Results summary"))
        elements.append(Spacer(1, 8)) # Gap after header line
        
        # Results summary table
        results_table = self._create_results_summary_table(embryos_data)
        elements.append(results_table)
        elements.append(Spacer(1, 8))
        
        # Results summary comment (optional, appears below table)
        results_summary_comment = self._clean(patient_data.get('results_summary_comment', ''))
        if results_summary_comment:
            comment_para = Paragraph(results_summary_comment, self.styles['PGTABodyText'])
            elements.append(comment_para)
            elements.append(Spacer(1, 8))
        
        elements.append(Spacer(1, 4))
        
        return elements
    
    def _clean(self, val, default=""):
        """Sanitize value to remove 'nan' and trim whitespace"""
        if val is None: return default
        s = str(val).strip()
        if s.lower() == "nan": return default
        return s
    
    def _wrap_text(self, text, bold=False, font_size=None, align='LEFT', max_width=None):
        """Wrap text in a Paragraph for table cells, with automatic Line Break support"""
        if not text: return ""
        
        # UI UX: Convert newlines (from Enter key) to PDF line breaks automatically
        content = str(text).replace('\r\n', '\n').replace('\r', '\n')
        content = content.replace('\n', '<br/>\u00A0') # Non-breaking space ensures line has height
        content = content.strip(' \t\r\f\v') # Strip horizontal whitespace
        
        # Robust 'nan' check
        if content.lower() == "nan" or content.lower() == "<br/>": # Skip if only nan or empty break
            if content.lower() != "<br/>": # Keep explicit breaks if user intentionally added them? 
                # Actually if it's JUST a break, they might want it. 
                # Let's only skip 'nan'.
                return "" if content.lower() == "nan" else Paragraph(content, self.styles['PGTALeftBodyText'])
            
        # Select appropriate style based on alignment
        if align == 'CENTER':
            style_name = 'PGTACenteredBodyText'
        elif align == 'LEFT':
            style_name = 'PGTALeftBodyText'  # Use LEFT aligned style to avoid justify gaps
        else:
            style_name = 'PGTABodyText'  # Default with justify
        
        # Determine style and font size override
        use_style = self.styles[style_name]
        if font_size:
            use_style = ParagraphStyle(
                name=f'{style_name}_custom_{font_size}',
                parent=self.styles[style_name],
                fontSize=font_size,
                leading=font_size * 1.2
            )
            
        # Apply explicit bold tag if requested (and not already in content)
        final_text = content
        if bold:
            # Avoid doubling bold tags if user/logic already added them
            if not (final_text.startswith('<b>') and final_text.endswith('</b>')):
                final_text = f"<b>{content}</b>"
            
        return Paragraph(final_text, use_style)

    def _wrap_label(self, text):
        """Wrap label text with forced RIGHT alignment and no word gaps"""
        if not text: return ""
        # Use <nobr> tags to prevent word breaking/justification
        return Paragraph(f"<nobr>{str(text)}</nobr>", self.styles['PGTALabelText'])

    def _create_patient_info_table(self, patient_data):
        """Create patient information table"""
        # Prepare data with Paragraph wrapping to prevent overlap
        # Standard widths for cover page: [108, 12, 131, 108, 12, 119] Total: 490pt
        
        # Patient name and spouse name - spouse on new line
        import re
        patient_name = re.sub(r'\s+', ' ', self._clean(patient_data.get('patient_name'))).strip()
        spouse_name = re.sub(r'\s+', ' ', self._clean(patient_data.get('spouse_name'))).strip()
        # Put spouse on new line with <br/> if present
        combined_name = f"{patient_name}<br/>{spouse_name}" if spouse_name else patient_name
        
        data = [
            [self._wrap_text('<b>PATIENT NAME</b>', True), self._wrap_text(':'), self._wrap_text(f"<b>{combined_name}</b>", max_width=140), self._wrap_text('<b>PIN</b>', True), self._wrap_text(':'), self._wrap_text(f"<b>{self._clean(patient_data.get('pin'))}</b>", max_width=144)],
            [self._wrap_text('<b>DATE OF BIRTH/ AGE</b>', True), self._wrap_text(':'), self._wrap_text(f"<b>{self._clean(patient_data.get('age'))}</b>", max_width=140), self._wrap_text('<b>SAMPLE NUMBER</b>', True), self._wrap_text(':'), self._wrap_text(f"<b>{self._clean(patient_data.get('sample_number'))}</b>", max_width=144)],
            [self._wrap_text('<b>REFERRING CLINICIAN</b>', True), self._wrap_text(':'), self._wrap_text(f"<b>{self._clean(patient_data.get('referring_clinician'))}</b>", max_width=140), self._wrap_text('<b>BIOPSY DATE</b>', True), self._wrap_text(':'), self._wrap_text(f"<b>{self._clean(patient_data.get('biopsy_date'))}</b>", max_width=144)],
            [self._wrap_text('<b>HOSPITAL/CLINIC</b>', True), self._wrap_text(':'), self._wrap_text(f"<b>{self._clean(patient_data.get('hospital_clinic'))}</b>", max_width=140), self._wrap_text('<b>SAMPLE COLLECTION DATE</b>', True), self._wrap_text(':'), self._wrap_text(f"<b>{self._clean(patient_data.get('sample_collection_date'))}</b>", max_width=144)],
            [self._wrap_text('<b>SPECIMEN</b>', True), self._wrap_text(':'), self._wrap_text(f"<b>{self._clean(patient_data.get('specimen'))}</b>", max_width=140), self._wrap_text('<b>SAMPLE RECEIPT DATE</b>', True), self._wrap_text(':'), self._wrap_text(f"<b>{self._clean(patient_data.get('sample_receipt_date'))}</b>", max_width=144)],
            [self._wrap_text('<b>BIOPSY PERFORMED BY</b>', True), self._wrap_text(':'), self._wrap_text(f"<b>{self._clean(patient_data.get('biopsy_performed_by'))}</b>", max_width=140), self._wrap_text('<b>REPORT DATE</b>', True), self._wrap_text(':'), self._wrap_text(f"<b>{self._clean(patient_data.get('report_date'))}</b>", max_width=144)]
        ]
        
        # Widths: left-label(108) + colon(12) + left-value(161) + right-label(108) + colon(12) + right-value(89) = 490pt
        # Right-side block starts at 281pt from left edge (col0+col1+col2)
        table = Table(data, colWidths=[108, 12, 161, 108, 12, 89], hAlign='LEFT')
        
        # Style table
        table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), self._get_font('SegoeUI-Bold', 'Helvetica-Bold')),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'), # Standard LEFT alignment
            ('ALIGN', (3, 0), (3, -1), 'LEFT'), # Standard LEFT alignment
            ('LEFTPADDING', (0, 0), (0, -1), 4),
            ('LEFTPADDING', (3, 0), (3, -1), 4),
            ('LEFTPADDING', (1, 0), (2, -1), 0),
            ('LEFTPADDING', (4, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(self.COLORS['patient_info_bg'])),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4), # Extra padding to fill box height
            ('TOPPADDING', (0, 0), (-1, -1), 4),
        ] + self._get_grid_style()))
        
        return table
    
    def _create_results_summary_table(self, embryos_data):
        """Create results summary table"""
        # Header row
        header_labels = ['S. No.', 'Sample', 'Result', 'MTcopy', 'Interpretation']
        data = [[self._wrap_text(label, bold=True, align='CENTER') for label in header_labels]]
        
        # Add embryo rows
        for idx, embryo in enumerate(embryos_data, 1):
            res_sum = self._clean(embryo.get('result_summary'))
            interp = self._clean(embryo.get('interpretation'))
            
            # Application of Red/Blue color logic
            res_color = self._get_result_color(res_sum, interp)
            
            # MTcopy: NA for non-euploid
            mtcopy = self._clean(embryo.get('mtcopy'), 'NA')
            if interp.upper() != "EUPLOID":
                mtcopy = "NA"
            
            data.append([
                self._wrap_text(str(idx), align='CENTER'),
                self._wrap_text(self._clean(embryo.get('embryo_id')), align='CENTER'),
                # Color only, no bold as per latest request
                self._wrap_text(self._wrap_colored(res_sum, res_color, bold=False), align='CENTER'),
                self._wrap_text(mtcopy, align='CENTER'),
                self._wrap_text(self._wrap_colored(interp, res_color, bold=False), align='CENTER')
            ])
        
        # Create table [Total: 496pt - Ensuring it fills the full content width]
        table = Table(data, colWidths=[50, 95, 185, 80, 86])
        
        # Style table
        table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, 0), self._get_font('Calibri-Bold', 'Helvetica-Bold')),
            ('FONTNAME', (0, 1), (-1, -1), self._get_font('Calibri', 'Helvetica')),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            # Header row - peach (exact from source)
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(self.COLORS['results_header_bg'])),
            # First column (S.No) - NO peach as per latest request (only header row)
            # All other data cells - light blue-grey (exact from source)
            # GRID updated: horizontal lines only (remove internal vertical lines)
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor(self.COLORS['patient_info_bg'])),
            ('LINEBELOW', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            # Explicitly ensure center alignment matching request for ALL data cells including S.No
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ] + self._get_grid_style()))
        
        return table
    
    def _build_methodology_page(self):
        """Build methodology and static content page - sections flow continuously"""
        elements = []
        
        # Methodology section - no KeepTogether, just natural flow
        elements.append(self._create_section_header("Methodology"))
        elements.append(Spacer(1, 8))
        elements.append(Paragraph(self.METHODOLOGY_TEXT, self.styles['PGTABodyText']))
        elements.append(Spacer(1, 12))
        
        # Mosaicism section - natural flow
        elements.append(self._create_section_header("Conditions for reporting mosaicism"))
        elements.append(Spacer(1, 8))
        elements.append(Paragraph(self.MOSAICISM_TEXT, self.styles['PGTABodyText']))
        elements.append(Spacer(1, 6))
        
        # Mosaicism bullets
        for bullet in self.MOSAICISM_BULLETS:
            elements.append(Paragraph(f"• {bullet}", self.styles['PGTABulletText']))
        elements.append(Spacer(1, 6))
        elements.append(Paragraph(self.MOSAICISM_CLINICAL, self.styles['PGTABodyText']))
        elements.append(Spacer(1, 12))
        
        # Limitations section - natural flow
        elements.append(self._create_section_header("Limitations"))
        elements.append(Spacer(1, 8))
        for limitation in self.LIMITATIONS:
            elements.append(Paragraph(f"• {limitation}", self.styles['PGTABulletText']))
        
        elements.append(Spacer(1, 12))
        elements.append(Spacer(1, 12))
        
        # References section - natural flow
        elements.append(self._create_section_header("References"))
        elements.append(Spacer(1, 8))
        for idx, ref in enumerate(self.REFERENCES, 1):
            elements.append(Paragraph(f"{idx}. {ref}", self.styles['PGTABodyText']))
        
        return elements
    
    def _build_embryo_page(self, patient_data, embryo_data):
        """Build individual embryo results page"""
        elements = []
        
        # Title repeated in embryo section as requested
        title = Paragraph(
            "Preimplantation Genetic Testing for Aneuploidies (PGT-A)",
            self.styles['PGTAReportTitle']
        )
        elements.append(title)
        elements.append(Spacer(1, 8))
        
        # Prepare info data with sanitation
        def _wrap_banner(text):
            if not text: return ""
            return Paragraph(str(text), self.styles['PGTABannerValueText'])

        # Patient name and spouse name - spouse on new line
        patient_name = self._clean(patient_data.get('patient_name'))
        spouse_name = self._clean(patient_data.get('spouse_name'))
        # Put spouse on new line with <br/> if present
        combined_name = f"{patient_name}<br/>{spouse_name}" if spouse_name else patient_name

        info_data = [[
            self._wrap_label('PATIENT NAME:'),
            _wrap_banner(f"<b>{combined_name}</b>"),
            Paragraph(f"<nobr>PIN:</nobr>", self.styles['PGTALabelTextRight']),
            _wrap_banner(f"<b>{self._clean(patient_data.get('pin'))}</b>")
        ]]
        
        # Optimized layout: Push PIN block right. PATIENT NAME (82pt), PIN (24pt)
        # Shifted slightly left (-30pt on spacer column): 82 + 254 + 24 + 130 = 490pt
        info_table = Table(info_data, colWidths=[82, 254, 24, 130], hAlign='LEFT')
        info_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(self.COLORS['patient_info_bg'])),
            ('FONTNAME', (0, 0), (-1, -1), self._get_font('SegoeUI-Bold', 'Helvetica-Bold')),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),  # Patient name label LEFT
            ('ALIGN', (2, 0), (2, -1), 'RIGHT'),  # PIN label RIGHT
            ('LEFTPADDING', (0, 0), (0, -1), 4),
            ('LEFTPADDING', (1, 0), (1, -1), 0),
            ('LEFTPADDING', (2, 0), (2, -1), 0),
            ('LEFTPADDING', (3, 0), (3, -1), 4),
            ('RIGHTPADDING', (0, 0), (0, -1), 0),
            ('RIGHTPADDING', (1, 0), (1, -1), 0),
            ('RIGHTPADDING', (2, 0), (2, -1), 4),
            ('RIGHTPADDING', (3, 0), (3, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ] + self._get_grid_style()))
        elements.append(info_table)
        elements.append(Spacer(1, 8))
        
        # PNDT Disclaimer in a grey box (Exact grey from source)
        disclaimer = Paragraph(
            "<b>This test does not reveal sex of the fetus & confers to PNDT act, 1994</b>",
            self.styles['PGTADisclaimer']
        )
        disclaimer_table = Table([[disclaimer]], colWidths=[490], hAlign='CENTER')
        disclaimer_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(self.COLORS['grey_bg'])),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ] + self._get_grid_style()))
        elements.append(KeepTogether(disclaimer_table))
        elements.append(Spacer(1, 12))
        
        # Application of Red/Blue color logic with sanitation
        res_text = self._clean(embryo_data.get('result_description'))
        autosomes_text = self._clean(embryo_data.get('autosomes'))
        interp_text = self._clean(embryo_data.get('interpretation'))
        sex_text = self._clean(embryo_data.get('sex_chromosomes', 'Normal'))
        result_summary_text = self._clean(embryo_data.get('result_summary', ''))
        
        # Color based on interpretation for interpretation field
        # Also pass result_summary so "Multiple chromosomal abnormalities" drives red color
        interp_color = self._get_result_color(result_summary_text, interp_text)
        
        # Autosomes Color Logic:
        # Blue (mosaic) = Has % sign (e.g., +15(~30%), -20(~51%), dup(9)...(~32%))
        # Red (non-mosaic) = del/dup/-/+ without %, or CNV status L/G/SL/SG
        # Red = Multiple chromosomal abnormalities (explicit result_summary check)
        # Black = Normal/Euploid
        auto_color = colors.black
        auto_upper = autosomes_text.upper()
        
        # Check for "Multiple chromosomal abnormalities" in result_summary first
        if "MULTIPLE CHROMOSOMAL ABNORMALITIES" in result_summary_text.upper():
            auto_color = colors.red
        # Check for Normal/Euploid
        elif 'NORMAL' in auto_upper or 'EUPLOID' in auto_upper or not autosomes_text.strip():
            auto_color = colors.black
        # Mosaic = has % sign
        elif '%' in autosomes_text:
            auto_color = colors.blue
        # Non-mosaic abnormalities (no % sign)
        elif any(x in auto_upper for x in ['DEL(', 'DUP(', '-', '+', 'STATUS L', 'STATUS G', 'STATUS SL', 'STATUS SG', ' SL', ' SG', ' L,', ' G,', ' L ', ' G ']) or auto_upper.endswith(' L') or auto_upper.endswith(' G'):
            auto_color = colors.red
        elif 'CNV STATUS' in auto_upper:
            auto_color = colors.red
        
        # Sex Chromosome Color
        sex_color = colors.black
        if "ABNORMAL" in sex_text.upper():
            sex_color = colors.red
        elif "MOSAIC" in sex_text.upper():
            sex_color = colors.blue

        # MTcopy: NA for non-euploid
        mtcopy = self._clean(embryo_data.get('mtcopy'), 'NA')
        if interp_text.upper() != "EUPLOID":
            mtcopy = "NA"
            
        # Embryo Identification matching source style
        # Font: Gill Sans MT,Bold, Size: 12.00
        # Use embryo_id_detail for detail pages, fallback to embryo_id if not present
        detail_embryo_id = self._clean(embryo_data.get('embryo_id_detail')) or self._clean(embryo_data.get('embryo_id'))
        embryo_id_style = ParagraphStyle(
            name='EmbryoIDStyle',
            parent=self.styles['Normal'],
            fontSize=12,
            leading=14,
            fontName=self._get_font('GillSansMT-Bold', 'Helvetica-Bold'),
            textColor=colors.HexColor(self.COLORS['blue_title'])
        )
        elements.append(Paragraph(f"<b>EMBRYO: {detail_embryo_id}</b>", embryo_id_style))
        elements.append(Spacer(1, 6))
        
        # Result row: color driven by result_summary (Multiple chromosomal abnormalities = red)
        detail_data = [
            [self._wrap_text(f"<b>Result:</b> {self._wrap_colored(res_text, colors.black, bold=False)}", False)],
            [self._wrap_text(f"<b>Autosomes:</b> {self._wrap_colored(autosomes_text, auto_color, bold=False)}", False)],
            [self._wrap_text(f"<b>Sex Chromosomes:</b> {self._wrap_colored(sex_text, sex_color, bold=False)}", False)],
            [self._wrap_text(f"<b>Interpretation:</b> {self._wrap_colored(interp_text, interp_color, bold=False)}", False)],
            [self._wrap_text(f"<b>MTcopy:</b> {mtcopy}", False)]
        ]
        
        # Summary table in detailed section [Total: 496pt]
        detail_table = Table(detail_data, colWidths=[490], hAlign='CENTER')
        detail_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(self.COLORS['patient_info_bg'])),
            ('LEFTPADDING', (0, 0), (-1, -1), 0), # Alignment fix
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ] + self._get_grid_style()))
        elements.append(detail_table)
        elements.append(Spacer(1, 12))
        
        # CNV Chart title - No line for this one
        elements.append(self._create_section_header("COPY NUMBER CHART", show_line=False))
        elements.append(Spacer(1, 6))
        
        # CNV Chart Image
        if 'cnv_image_path' in embryo_data and embryo_data['cnv_image_path'] and os.path.exists(embryo_data['cnv_image_path']):
            try:
                # Add image, keeping aspect ratio but fitting within width
                img = Image(embryo_data['cnv_image_path'], width=self.CONTENT_WIDTH)
                
                # Adjust height proportionally
                aspect = img.imageWidth / img.imageHeight
                img.drawHeight = self.CONTENT_WIDTH / aspect
                
                img.hAlign = 'CENTER'
                elements.append(img)
                elements.append(Spacer(1, 12))
            except Exception as e:
                print(f"Error loading image: {e}")
        
        # CNV table - Skip for Inconclusive results (only skip table, not chart)
        result_summary = self._clean(embryo_data.get('result_summary', ''))
        result_desc = self._clean(embryo_data.get('result_description', ''))
        is_inconclusive = "INCONCLUSIVE" in result_summary.upper() or "INCONCLUSIVE" in result_desc.upper() or "INCONCLUSIVE" in interp_text.upper()
        
        # Add inconclusive comment under CNV chart if present
        if is_inconclusive:
            inconclusive_comment = self._clean(embryo_data.get('inconclusive_comment', ''))
            if inconclusive_comment:
                comment_para = Paragraph(
                    f"{inconclusive_comment}",
                    self.styles['PGTABodyText']
                )
                elements.append(comment_para)
                elements.append(Spacer(1, 12))
        
        if not is_inconclusive:
            cnv_table = self._create_cnv_table(embryo_data)
            elements.append(cnv_table)
            elements.append(Spacer(1, 6))
            
            # Legend
            legend = Paragraph(
                "<i>N – Normal, G-Gain, L-Loss, SG-Segmental Gain, SL-Segmental Loss, "
                "M-Mosaic, MG- Mosaic Gain, ML-Mosaic Loss, SMG-Segmental Mosaic Gain, "
                "SML-Segmental Mosaic Loss</i>",
                self.styles['PGTASmallText']
            )
            elements.append(legend)
            elements.append(Spacer(1, 12))
        
        return elements
    
    def _create_cnv_table(self, embryo_data):
        """Create CNV status table"""
        # Get chromosome statuses
        chr_statuses = embryo_data.get('chromosome_statuses', {})
        mosaic_percentages = embryo_data.get('mosaic_percentages', {})
        
        autosomes = str(embryo_data.get('autosomes', '')).upper()
        sex_chrs = str(embryo_data.get('sex_chromosomes', '')).upper()
        
        # Check for actual mosaic percentage values (not empty, not dash, must contain a digit)
        import re as re_mos
        has_mosaic = any(
            v and str(v).strip() and str(v).strip() != '-' and re_mos.search(r'\d', str(v))
            for v in mosaic_percentages.values()
        )
        
        # Rule: If Autosomes is Normal/Euploid & Sex chromosome is Mosaic gain/loss -> omit Mosaic(%) row
        is_autosomes_normal = 'NORMAL' in autosomes or 'EUPLOID' in autosomes or not autosomes.strip()
        is_sex_mosaic = 'MOSAIC' in sex_chrs
        
        if is_autosomes_normal and is_sex_mosaic:
            has_mosaic = False
            
        if has_mosaic:
            # Header set to 9pt
            header = [self._wrap_text('Chromosome', bold=True, align='CENTER', font_size=9)] + [self._wrap_text(str(i), bold=True, align='CENTER', font_size=9) for i in range(1, 23)]
            cnv_row = [self._wrap_text('CNV status', bold=True, align='CENTER', font_size=9)]
            mosaic_row = [self._wrap_text('Mosaic (%)', bold=True, align='CENTER', font_size=9)]
            
            for i in range(1, 23):
                status = chr_statuses.get(str(i), 'N')
                perc = mosaic_percentages.get(str(i), '-')
                s_color = self._get_status_color(status)
                
                # Dynamic font sizing for balancing visibility (Normal: 9pt, SMG/SML: 8pt)
                f_size = 8 if len(status) > 2 else 9
                cnv_row.append(self._wrap_text(self._wrap_colored(status, s_color, bold=True), bold=True, font_size=f_size, align='CENTER'))
                mosaic_row.append(self._wrap_text(self._wrap_colored(str(perc), s_color, bold=True), bold=True, font_size=9, align='CENTER'))
            
            data = [header, cnv_row, mosaic_row]
            # Final optimized width: "Chromosome" widened to 75pt to ensure NO wrap. 
            # Remaining width (496 - 75 = 421) / 22 columns = ~19.13pt per data column
            col_widths = [75] + [19.13] * 22
        else:
            header = [self._wrap_text('Chromosome', bold=True, align='CENTER', font_size=9)] + [self._wrap_text(str(i), bold=True, align='CENTER', font_size=9) for i in range(1, 23)]
            cnv_row = [self._wrap_text('CNV status', bold=True, align='CENTER', font_size=9)]
            for i in range(1, 23):
                status = chr_statuses.get(str(i), 'N')
                s_color = self._get_status_color(status)
                
                # Dynamic font sizing for balancing visibility (Normal: 9pt, SMG/SML: 8pt)
                f_size = 8 if len(status) > 2 else 9
                cnv_row.append(self._wrap_text(self._wrap_colored(status, s_color, bold=True), bold=True, font_size=f_size, align='CENTER'))
                
            data = [header, cnv_row]
            col_widths = [75] + [19.13] * 22
        
        # Create table
        table = Table(data, colWidths=col_widths)
        
        # Style table
        table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), self._get_font('SegoeUI-Bold', 'Helvetica-Bold')),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            # All cells - light blue-grey (exact from source) as requested
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(self.COLORS['patient_info_bg'])),
            # White grid lines for "white line gaps"
            # White grid lines for "white line gaps"
            ('GRID', (0, 0), (-1, -1), 1.0, colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            # Zero padding for data columns to maximize horizontal space for 8pt/9pt fonts
            ('LEFTPADDING', (1, 0), (-1, -1), 0),
            ('RIGHTPADDING', (1, 0), (-1, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ] + self._get_grid_style()))
        
        return table

    def _create_signature_table(self):
        """Create signature section with precise structural metrics from source PDF"""
        elements = []
        
        # Leading text: exact color #4F81BD, size 12.48pt
        elements.append(Paragraph(
            "<b>This report has been reviewed and approved by: </b>", 
            self.styles['PGTASigApproval']
        ))
        # Exact source vertical gap (12.7pt from text baseline to image top)
        elements.append(Spacer(1, 12.7))
        
        # Individual Base64 images for signatures as requested
        try:
            from io import BytesIO
            def get_sig_img(b64):
                if not b64: return Paragraph('', self.styles['Normal'])
                img_data = base64.b64decode(b64)
                return Image(BytesIO(img_data), width=100, height=40)

            sig1 = get_sig_img(SIGN_ANAND_B64)
            sig2 = get_sig_img(SIGN_SACHIN_B64)
            sig3 = get_sig_img(SIGN_DIRECTOR_B64)
            
            sig_img_table = Table([[sig1, sig2, sig3]], colWidths=[156, 156, 156])
            sig_img_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'BOTTOM'),
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ] + self._get_grid_style()))
            elements.append(sig_img_table)
        except Exception as e:
            print(f"Error drawing individual signatures: {e}")

        # Names and Titles below signatures - Using SegoeUI Normal weight
        data = []
        names_row = []
        titles_row = []
        
        sig_name_style = ParagraphStyle('SigName', parent=self.styles['Normal'], 
                                       fontName=self._get_font('SegoeUI', 'Helvetica'), 
                                       fontSize=11.04, alignment=TA_CENTER)
        sig_title_style = ParagraphStyle('SigTitle', parent=self.styles['Normal'], 
                                        fontName=self._get_font('SegoeUI', 'Helvetica'), 
                                        fontSize=11.04, alignment=TA_CENTER)

        for sig in self.SIGNATURES:
            names_row.append(Paragraph(sig['name'], sig_name_style))
            titles_row.append(Paragraph(sig['title'], sig_title_style))
        
        data = [names_row, titles_row]
        
        # Table width 468 [3 x 156]
        table = Table(data, colWidths=[156, 156, 156])
        table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('GRID', (0, 0), (-1, -1), 0, colors.white),
            ('BOX', (0, 0), (-1, -1), 0, colors.white),
            ('LINEABOVE', (0, 0), (-1, -1), 0, colors.white),
            ('LINEBELOW', (0, 0), (-1, -1), 0, colors.white),
            ('LINEBEFORE', (0, 0), (-1, -1), 0, colors.white),
            ('LINEAFTER', (0, 0), (-1, -1), 0, colors.white),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ] + self._get_grid_style()))
        
        elements.append(table)
        
        return KeepTogether(elements)

    def _create_section_header(self, text, show_line=True):
        """Create a section header with navy blue text and a slight lighter line below"""
        header = Paragraph(f"<b>{text}</b>", self.styles["PGTASectionHeader"])
        
        if not show_line:
            return KeepTogether([header])
            
        # Create a single-cell table for the line
        header_table = Table([[header]], colWidths=[490], hAlign='CENTER')
        header_table.setStyle(TableStyle([
            # Slight grey line, lighter than full black
            ("LINEBELOW", (0, 0), (-1, -1), 0.5, colors.HexColor("#989998")),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            # Increased bottom padding to 6pt as requested for visual gap between text and line
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ] + self._get_grid_style()))
        return KeepTogether(header_table)

    def _get_result_color(self, result_text, interpretation_text):
        """Determine if text should be Red (Aneuploid), Blue (Mosaic) or Black (Euploid)"""
        res_up = result_text.upper() if result_text else ""
        int_up = interpretation_text.upper() if interpretation_text else ""
        
        # Euploid = Black (check first for explicit euploid)
        if "EUPLOID" in int_up and "ANEUPLOID" not in int_up:
            return colors.black
        
        # Red Logic - Aneuploid and related abnormalities
        red_keywords = ["MONOSOMY", "TRISOMY", "SEGMENTAL GAIN", "SEGMENTAL LOSS", 
                        "MULTIPLE CHROMOSOMAL ABNORMALITIES", "ANEUPLOID", "CHAOTIC EMBRYO"]
        if any(kw in res_up for kw in red_keywords) or any(kw in int_up for kw in red_keywords):
            return colors.red
            
        # Blue Logic - Mosaic
        blue_keywords = ["MOSAIC CHROMOSOME COMPLEMENT", "LOW LEVEL MOSAIC", 
                         "HIGH LEVEL MOSAIC", "COMPLEX MOSAIC", "MULTIPLE MOSAIC"]
        if any(kw in res_up for kw in blue_keywords) or any(kw in int_up for kw in blue_keywords):
            return colors.blue
            
        return colors.black

    def _get_autosome_color(self, autosome_text):
        """Special color logic for autosomes field"""
        if not autosome_text: return colors.black
        txt = autosome_text.upper()
        # Red: Multiple chromosomal abnormalities
        if "MULTIPLE CHROMOSOMAL ABNORMALITIES" in txt:
            return colors.red
        # Blue: Mosaic variants
        if "MULTIPLE MOSAIC CHROMOSOME COMPLEMENT" in txt:
            return colors.blue
        return colors.black

    def _get_status_color(self, status):
        """Color logic for CNV status codes"""
        if not status: return colors.black
        s = status.upper().strip()
        
        # Combination codes (must check before single codes)
        red_combos = ["SL/SG", "SG/SL"]
        blue_combos = ["SML/SMG", "SMG/SML"]
        if s in red_combos: return colors.red
        if s in blue_combos: return colors.blue
        
        # Single codes
        red_codes = ["G", "L", "SG", "SL"]
        blue_codes = ["M", "MG", "ML", "SMG", "SML"]
        if s in red_codes: return colors.red
        if s in blue_codes: return colors.blue
        
        # Numeric check for mosaic percentage
        try:
            val = float(s.replace('%', ''))
            if val > 0: return colors.blue
        except:
            pass
            
        return colors.black

    def _wrap_colored(self, text, color, bold=False):
        """Standard wrapper for colored text with optional bolding"""
        if not text: return text
        if color == colors.black:
            return f"<b>{text}</b>" if bold else str(text)
        hex_color = color.hexval()[2:] # Remove 0x
        if len(hex_color) > 6: hex_color = hex_color[2:] # Handle alpha
        
        if bold:
            return f'<b><font color="#{hex_color}">{text}</font></b>'
        return f'<font color="#{hex_color}">{text}</font>'


if __name__ == "__main__":
    # Test the template
    template = PGTAReportTemplate()
    
    # Sample data
    patient_data = {
        'patient_name': 'Mrs. Priya (PNM00791)',
        'patient_spouse': 'Mrs. Priya (PNM00791)',
        'spouse_name': 'Mr. Saranraj',
        'pin': 'AND25630004206',
        'age': '34 Years',
        'sample_number': '632504349',
        'referring_clinician': 'Dr. Ajantha. B',
        'biopsy_date': '03-01-2026',
        'hospital_clinic': 'Rhea Healthcare Private Limited Annanagar (NOVA IVF)',
        'sample_collection_date': '03-01-2026',
        'specimen': 'DAY 5 TROPHECTODERM BIOPSY',
        'sample_receipt_date': '03-01-2026',
        'biopsy_performed_by': 'Raj Priya Pandian',
        'report_date': '14-01-2026',
        'indication': 'History of implantation failure.'
    }
    
    embryos_data = [
        {
            'embryo_id': 'PS4',
            'result_summary': 'Trisomy of chromosome 16',
            'mtcopy': 'NA',
            'interpretation': 'Aneuploid',
            'result_description': 'The embryo contains abnormal chromosome complement',
            'autosomes': 'Trisomy of chromosome 16',
            'sex_chromosomes': 'Normal',
            'chromosome_statuses': {str(i): 'N' for i in range(1, 23)},
            'mosaic_percentages': {},
            'cnv_image_path': os.path.join(os.path.dirname(os.path.abspath(__file__)), "PRIYA-PS4_L00_R1_noXY_nomos.png")
        }
    ]
    
    # Update chromosome 16 to Segmental Mosaic Loss (SML) to test font scaling
    embryos_data[0]['chromosome_statuses']['16'] = 'SML'
    
    # Generate PDF
    output_path = "test_report.pdf"
    template.generate_pdf(output_path, patient_data, embryos_data, show_grid=True)
    print(f"Test report generated: {output_path}")
