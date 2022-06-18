from pathlib import Path

'''
Some sample XML snippets!

xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"

# theme/theme1.xml
        <a:fontScheme name="Pretendard">
            <a:majorFont>
                <a:latin typeface="Pretendard"/>
                <a:ea typeface="Pretendard"/>
                <a:cs typeface="Pretendard"/>
                <a:font script="Hang" typeface="Pretendard"/>
            </a:majorFont>
            <a:minorFont>
                <a:latin typeface="Pretendard"/>
                <a:ea typeface="Pretendard"/>
                <a:cs typeface="Pretendard"/>
                <a:font script="Hang" typeface="Pretendard"/>
            </a:minorFont>
        </a:fontScheme>

# presentation.xml
    <p:defaultTextStyle>
        <a:defPPr>
            <a:defRPr lang="ko-Kore-KR"/>
        </a:defPPr>
        <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
            <a:defRPr sz="1800" kern="1200">
                <a:solidFill>
                    <a:schemeClr val="tx1"/>
                </a:solidFill>
                <a:latin typeface="+mn-lt"/>
                <a:ea typeface="+mn-ea"/>
                <a:cs typeface="+mn-cs"/>
            </a:defRPr>
        </a:lvl1pPr>
        ...
    </p:defaultTextStyle>

# slideMasters/slideMaster1.xml
        <p:titleStyle>
            <a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:lnSpc>
                    <a:spcPct val="90000"/>
                </a:lnSpc>
                <a:spcBef>
                    <a:spcPct val="0"/>
                </a:spcBef>
                <a:buNone/>
                <a:defRPr sz="4400" b="1" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1"/>
                    </a:solidFill>
                    <a:latin typeface="+mj-lt"/>
                    <a:ea typeface="+mj-ea"/>
                    <a:cs typeface="+mj-cs"/>
                </a:defRPr>
            </a:lvl1pPr>
        </p:titleStyle>
        <p:bodyStyle>
            <a:lvl1pPr marL="228600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:lnSpc>
                    <a:spcPct val="90000"/>
                </a:lnSpc>
                <a:spcBef>
                    <a:spcPts val="1000"/>
                </a:spcBef>
                <a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
                <a:buChar char="•"/>
                <a:defRPr sz="2800" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1"/>
                    </a:solidFill>
                    <a:latin typeface="+mn-lt"/>
                    <a:ea typeface="+mn-ea"/>
                    <a:cs typeface="+mn-cs"/>
                </a:defRPr>
            </a:lvl1pPr>
            ...
        </p:bodyStyle>

# slideLayouts/slideLayout1.xml
        (TODO)


# slides/slide1.xml
# slides/slide2.xml
# slides/slide3.xml
'''

def fix_theme_font(work_path: Path, major_font: str, minor_font: str) -> None:
    theme_path = work_path / 'ppt' / 'theme' / 'theme1.xml'
    # TODO: themeN.xml: change majorFont and minorFont in fontScheme


def normalize_master_fonts(work_path: Path, bullet_font: str = "+mn-lt") -> None:
    # TODO: presentation.xml: change defRPr typefaces in defaultTextStyle
    # TODO: slideMasterN.xml: change buFont typeface
    # TODO: slideMasterN.xml: change defRPr typefaces in titleStyle and bodyStyle
    pass


def normalize_slide_fonts(work_path: Path) -> None:
    # TODO: ...
    pass
