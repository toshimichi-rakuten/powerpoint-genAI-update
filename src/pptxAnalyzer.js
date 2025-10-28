/**
 * ファイル名: src/pptxAnalyzer.js
 * 説明:
 *   既存のPPTXファイルを解析してAI用のプロンプト付きJSON形式で出力するモジュール。
 *   powerpoint-updateのpopup.jsから完全な解析ロジックを統合。
 *
 * 主な機能:
 *   - PPTXファイルの読み込みとZIP展開
 *   - テーマカラーの抽出と変換
 *   - スライド要素（テキスト、図形、表、線）の完全解析
 *   - テンプレート情報（背景色、スライド番号、固定画像）の抽出
 *   - マスタースタイルの解析
 *   - 箇条書きの解析
 *   - PptxGenJS用の詳細プロンプト付きJSON生成
 */

// 単位変換関数
function emuToInch(emu) {
  return emu / 914400;
}

function emuToPoint(emu) {
  return emu / 12700;
}

function fontSizeToPoint(szValue) {
  return szValue / 100;
}

function normalizeColorHex(hex) {
  if (!hex) return '';
  hex = hex.replace(/^#/, '');
  return hex.toUpperCase();
}

// RGB値にlumModとlumOffを適用する関数
function applyLuminanceModifiers(rgbHex, lumMod, lumOff) {
  if (!rgbHex) return '';

  const r = parseInt(rgbHex.substring(0, 2), 16);
  const g = parseInt(rgbHex.substring(2, 4), 16);
  const b = parseInt(rgbHex.substring(4, 6), 16);

  let newR = r * (lumMod / 100);
  let newG = g * (lumMod / 100);
  let newB = b * (lumMod / 100);

  if (lumOff !== null && lumOff !== 0) {
    const offset = 255 * (lumOff / 100);
    newR += offset;
    newG += offset;
    newB += offset;
  }

  newR = Math.max(0, Math.min(255, Math.round(newR)));
  newG = Math.max(0, Math.min(255, Math.round(newG)));
  newB = Math.max(0, Math.min(255, Math.round(newB)));

  const toHex = (n) => n.toString(16).padStart(2, '0').toUpperCase();
  return toHex(newR) + toHex(newG) + toHex(newB);
}

// schemeカラーをRGB値に変換する関数
function resolveSchemeColor(schemeColorName, lumMod, lumOff, themeColors) {
  if (!themeColors || !schemeColorName) return '';

  const baseColor = themeColors[schemeColorName];
  if (!baseColor) {
    console.log(`テーマカラー ${schemeColorName} が見つかりません`);
    return '';
  }

  return applyLuminanceModifiers(baseColor, lumMod || 100, lumOff || 0);
}

// テーマカラーを読み込む関数
async function loadThemeColors(zip) {
  try {
    const themeFile = zip.file('ppt/theme/theme1.xml');
    if (!themeFile) {
      console.log('テーマファイルが見つかりません');
      return null;
    }

    const themeXml = await themeFile.async('string');
    const parser = new DOMParser();
    const doc = parser.parseFromString(themeXml, 'application/xml');

    const themeColors = {};

    const clrScheme = Array.from(doc.getElementsByTagName('*')).find(el =>
      el.tagName.endsWith(':clrScheme') || el.localName === 'clrScheme'
    );

    if (!clrScheme) {
      console.log('カラースキームが見つかりません');
      return null;
    }

    const colorElements = [
      'dk1', 'lt1', 'dk2', 'lt2',
      'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
      'hlink', 'folHlink'
    ];

    colorElements.forEach(colorName => {
      const colorEl = Array.from(clrScheme.getElementsByTagName('*')).find(el =>
        el.tagName.endsWith(`:${colorName}`) || el.localName === colorName
      );

      if (colorEl) {
        const srgbClr = Array.from(colorEl.getElementsByTagName('*')).find(el =>
          el.tagName.endsWith(':srgbClr') || el.localName === 'srgbClr'
        );
        const sysClr = Array.from(colorEl.getElementsByTagName('*')).find(el =>
          el.tagName.endsWith(':sysClr') || el.localName === 'sysClr'
        );

        if (srgbClr) {
          themeColors[colorName] = normalizeColorHex(srgbClr.getAttribute('val'));
        } else if (sysClr) {
          themeColors[colorName] = normalizeColorHex(sysClr.getAttribute('lastClr'));
        }
      }
    });

    themeColors['bg1'] = themeColors['lt1'] || 'FFFFFF';
    themeColors['bg2'] = themeColors['lt2'] || 'E7E6E6';
    themeColors['tx1'] = themeColors['dk1'] || '000000';
    themeColors['tx2'] = themeColors['dk2'] || '44546A';

    console.log('テーマカラーを読み込みました:', themeColors);
    return themeColors;
  } catch (err) {
    console.log('テーマカラーの読み込みエラー:', err.message);
    return null;
  }
}

// 色情報を抽出（srgbClrまたはschemeClr対応、RGB値に変換）
function extractColor(element, themeColors) {
  if (!element) return '';

  const srgbClr = Array.from(element.getElementsByTagName("*")).find(el =>
    el.tagName.endsWith(":srgbClr") || el.localName === "srgbClr"
  );

  if (srgbClr) {
    const val = srgbClr.getAttribute('val');
    return normalizeColorHex(val || '');
  }

  const schemeClr = Array.from(element.getElementsByTagName("*")).find(el =>
    el.tagName.endsWith(":schemeClr") || el.localName === "schemeClr"
  );

  if (schemeClr) {
    const schemeName = schemeClr.getAttribute('val');
    const lumModEl = Array.from(schemeClr.getElementsByTagName("*")).find(el =>
      el.tagName.endsWith(":lumMod") || el.localName === "lumMod"
    );
    const lumOffEl = Array.from(schemeClr.getElementsByTagName("*")).find(el =>
      el.tagName.endsWith(":lumOff") || el.localName === "lumOff"
    );

    const lumMod = lumModEl ? parseInt(lumModEl.getAttribute('val') || '100000') / 1000 : 100;
    const lumOff = lumOffEl ? parseInt(lumOffEl.getAttribute('val') || '0') / 1000 : 0;

    return resolveSchemeColor(schemeName, lumMod, lumOff, themeColors);
  }

  return '';
}

// 罫線情報を抽出
function extractBorderInfo(lnElement, themeColors) {
  if (!lnElement) return { width: 0, color: '', dashType: 'solid' };

  const borderInfo = {
    width: parseInt(lnElement.getAttribute('w') || '0'),
    color: '',
    dashType: 'solid'
  };

  const noFill = Array.from(lnElement.getElementsByTagName("*")).find(el =>
    el.tagName.endsWith(":noFill") || el.localName === "noFill"
  );

  if (noFill) {
    borderInfo.width = 0;
    return borderInfo;
  }

  const solidFill = Array.from(lnElement.getElementsByTagName("*")).find(el =>
    el.tagName.endsWith(":solidFill") || el.localName === "solidFill"
  );

  if (solidFill) {
    borderInfo.color = extractColor(solidFill, themeColors);
  }

  const prstDash = Array.from(lnElement.getElementsByTagName("*")).find(el =>
    el.tagName.endsWith(":prstDash") || el.localName === "prstDash"
  );

  if (prstDash) {
    borderInfo.dashType = prstDash.getAttribute('val') || 'solid';
  }

  return borderInfo;
}

// テンプレート情報を抽出
async function extractTemplateInfo(zip, slidePath, themeColors) {
  try {
    const template = {
      background: "",
      defaultTextColor: "",
      slideNumber: null,
      fixedImages: []
    };

    const slideNum = slidePath.match(/slide(\d+)\.xml/)[1];
    const slideRelsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`;
    const slideRelsFile = zip.file(slideRelsPath);

    if (!slideRelsFile) {
      console.log('スライドの関係ファイルが見つかりません');
      return template;
    }

    const slideRelsXml = await slideRelsFile.async('string');
    const slideRelsDoc = new DOMParser().parseFromString(slideRelsXml, 'application/xml');

    const layoutRel = Array.from(slideRelsDoc.getElementsByTagName('Relationship')).find(rel =>
      rel.getAttribute('Type').includes('slideLayout')
    );

    if (!layoutRel) {
      console.log('スライドレイアウトの参照が見つかりません');
      return template;
    }

    const layoutPath = `ppt/slideLayouts/${layoutRel.getAttribute('Target').split('/').pop()}`;

    const layoutRelsPath = layoutPath.replace('.xml', '.xml.rels').replace('slideLayouts/', 'slideLayouts/_rels/');
    const layoutRelsFile = zip.file(layoutRelsPath);

    if (!layoutRelsFile) {
      console.log('レイアウトの関係ファイルが見つかりません');
      return template;
    }

    const layoutRelsXml = await layoutRelsFile.async('string');
    const layoutRelsDoc = new DOMParser().parseFromString(layoutRelsXml, 'application/xml');

    const masterRel = Array.from(layoutRelsDoc.getElementsByTagName('Relationship')).find(rel =>
      rel.getAttribute('Type').includes('slideMaster')
    );

    if (!masterRel) {
      console.log('スライドマスターの参照が見つかりません');
      return template;
    }

    const masterPath = `ppt/slideMasters/${masterRel.getAttribute('Target').split('/').pop()}`;

    // スライドマスターから背景色を抽出
    const masterFile = zip.file(masterPath);
    if (masterFile) {
      const masterXml = await masterFile.async('string');
      const masterDoc = new DOMParser().parseFromString(masterXml, 'application/xml');

      const bg = Array.from(masterDoc.getElementsByTagName('*')).find(el =>
        el.tagName.endsWith(':bg') || el.localName === 'bg'
      );

      if (bg) {
        const solidFill = Array.from(bg.getElementsByTagName('*')).find(el =>
          el.tagName.endsWith(':solidFill') || el.localName === 'solidFill'
        );

        if (solidFill) {
          template.background = extractColor(solidFill, themeColors);
          console.log(`背景色: ${template.background}`);
        }
      }
    }

    // スライドレイアウトから固定要素を抽出
    const layoutFile = zip.file(layoutPath);
    if (layoutFile) {
      const layoutXml = await layoutFile.async('string');
      const layoutDoc = new DOMParser().parseFromString(layoutXml, 'application/xml');

      // スライド番号を抽出
      const slideNumField = Array.from(layoutDoc.getElementsByTagName('*')).find(el =>
        (el.tagName.endsWith(':fld') || el.localName === 'fld') &&
        el.getAttribute('type') === 'slidenum'
      );

      if (slideNumField) {
        // スライド番号の位置とスタイルを取得
        const sp = slideNumField.closest('p\\:sp, sp');
        if (sp) {
          const xfrm = Array.from(sp.getElementsByTagName('*')).find(el =>
            el.tagName.endsWith(':xfrm') || el.localName === 'xfrm'
          );

          if (xfrm) {
            const off = Array.from(xfrm.getElementsByTagName('*')).find(el =>
              el.tagName.endsWith(':off') || el.localName === 'off'
            );
            const ext = Array.from(xfrm.getElementsByTagName('*')).find(el =>
              el.tagName.endsWith(':ext') || el.localName === 'ext'
            );

            if (off && ext) {
              const x = parseInt(off.getAttribute('x') || '0');
              const y = parseInt(off.getAttribute('y') || '0');
              const w = parseInt(ext.getAttribute('cx') || '0');
              const h = parseInt(ext.getAttribute('cy') || '0');

              // スタイル情報を取得
              const rPr = Array.from(slideNumField.getElementsByTagName('*')).find(el =>
                el.tagName.endsWith(':rPr') || el.localName === 'rPr'
              );

              template.slideNumber = {
                x: parseFloat(emuToInch(x).toFixed(3)),
                y: parseFloat(emuToInch(y).toFixed(3)),
                w: parseFloat(emuToInch(w).toFixed(3)),
                h: parseFloat(emuToInch(h).toFixed(3)),
                fontSize: 9,
                font: "",
                color: "000000",
                bold: false,
                align: "right"
              };

              if (rPr) {
                const sz = rPr.getAttribute('sz');
                if (sz) template.slideNumber.fontSize = parseFloat(fontSizeToPoint(parseInt(sz)).toFixed(1));

                const b = rPr.getAttribute('b');
                if (b) template.slideNumber.bold = b === '1';

                const latin = Array.from(rPr.getElementsByTagName('*')).find(el =>
                  el.tagName.endsWith(':latin') || el.localName === 'latin'
                );
                if (latin) template.slideNumber.font = latin.getAttribute('typeface') || '';

                const solidFill = Array.from(rPr.getElementsByTagName('*')).find(el =>
                  el.tagName.endsWith(':solidFill') || el.localName === 'solidFill'
                );
                if (solidFill) {
                  template.slideNumber.color = extractColor(solidFill, themeColors);
                }
              }

              // テキスト配置を取得
              const pPr = Array.from(sp.getElementsByTagName('*')).find(el =>
                el.tagName.endsWith(':pPr') || el.localName === 'pPr'
              );
              if (pPr) {
                const algn = pPr.getAttribute('algn');
                if (algn) template.slideNumber.align = algn;
              }

              console.log(`スライド番号: x=${template.slideNumber.x}, y=${template.slideNumber.y}, size=${template.slideNumber.fontSize}pt`);
            }
          }
        }
      }

      // 固定画像を抽出 (userDrawn="1"の画像)
      const pics = Array.from(layoutDoc.getElementsByTagName('*')).filter(el =>
        (el.tagName.endsWith(':pic') || el.localName === 'pic')
      );

      for (const pic of pics) {
        const nvPr = Array.from(pic.getElementsByTagName('*')).find(el =>
          el.tagName.endsWith(':nvPr') || el.localName === 'nvPr'
        );

        // userDrawn="1"の画像のみ抽出
        if (nvPr && nvPr.getAttribute('userDrawn') === '1') {
          const xfrm = Array.from(pic.getElementsByTagName('*')).find(el =>
            el.tagName.endsWith(':xfrm') || el.localName === 'xfrm'
          );

          if (xfrm) {
            const off = Array.from(xfrm.getElementsByTagName('*')).find(el =>
              el.tagName.endsWith(':off') || el.localName === 'off'
            );
            const ext = Array.from(xfrm.getElementsByTagName('*')).find(el =>
              el.tagName.endsWith(':ext') || el.localName === 'ext'
            );

            if (off && ext) {
              const x = parseInt(off.getAttribute('x') || '0');
              const y = parseInt(off.getAttribute('y') || '0');
              const w = parseInt(ext.getAttribute('cx') || '0');
              const h = parseInt(ext.getAttribute('cy') || '0');

              const cNvPr = Array.from(pic.getElementsByTagName('*')).find(el =>
                el.tagName.endsWith(':cNvPr') || el.localName === 'cNvPr'
              );
              const name = cNvPr ? cNvPr.getAttribute('name') : 'image';

              template.fixedImages.push({
                name: name,
                x: parseFloat(emuToInch(x).toFixed(3)),
                y: parseFloat(emuToInch(y).toFixed(3)),
                w: parseFloat(emuToInch(w).toFixed(3)),
                h: parseFloat(emuToInch(h).toFixed(3))
              });

              console.log(`固定画像: ${name} at (${template.fixedImages[template.fixedImages.length-1].x}, ${template.fixedImages[template.fixedImages.length-1].y})`);
            }
          }
        }
      }
    }

    return template;
  } catch (err) {
    console.log('テンプレート情報の抽出エラー:', err.message);
    return {
      background: "",
      defaultTextColor: "",
      slideNumber: null,
      fixedImages: []
    };
  }
}

// スライドマスターからスタイル情報を抽出
async function extractMasterStyles(zip, slidePath) {
  try {
    const masterStyles = {
      titleStyle: null,
      bodyStyle: null
    };

    const slideNum = slidePath.match(/slide(\d+)\.xml/)[1];
    const slideRelsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`;
    const slideRelsFile = zip.file(slideRelsPath);

    if (!slideRelsFile) {
      return masterStyles;
    }

    const slideRelsXml = await slideRelsFile.async('string');
    const slideRelsDoc = new DOMParser().parseFromString(slideRelsXml, 'application/xml');

    const layoutRel = Array.from(slideRelsDoc.getElementsByTagName('Relationship')).find(rel =>
      rel.getAttribute('Type').includes('slideLayout')
    );

    if (!layoutRel) {
      return masterStyles;
    }

    const layoutPath = `ppt/${layoutRel.getAttribute('Target').replace('../', '')}`;

    const layoutRelsPath = layoutPath.replace('.xml', '.xml.rels').replace('slideLayouts/', 'slideLayouts/_rels/');
    const layoutRelsFile = zip.file(layoutRelsPath);

    if (!layoutRelsFile) {
      return masterStyles;
    }

    const layoutRelsXml = await layoutRelsFile.async('string');
    const layoutRelsDoc = new DOMParser().parseFromString(layoutRelsXml, 'application/xml');

    const masterRel = Array.from(layoutRelsDoc.getElementsByTagName('Relationship')).find(rel =>
      rel.getAttribute('Type').includes('slideMaster')
    );

    if (!masterRel) {
      return masterStyles;
    }

    const masterPath = `ppt/${masterRel.getAttribute('Target').replace('../', '')}`;

    const masterFile = zip.file(masterPath);
    if (!masterFile) {
      return masterStyles;
    }

    const masterXml = await masterFile.async('string');
    const masterDoc = new DOMParser().parseFromString(masterXml, 'application/xml');

    // タイトルスタイルを抽出
    const titleStyle = Array.from(masterDoc.getElementsByTagName("*")).find(el =>
      el.tagName.endsWith(":titleStyle") || el.localName === "titleStyle"
    );

    if (titleStyle) {
      const lvl1pPr = Array.from(titleStyle.getElementsByTagName("*")).find(el =>
        el.tagName.endsWith(":lvl1pPr") || el.localName === "lvl1pPr"
      );

      if (lvl1pPr) {
        const defRPr = Array.from(lvl1pPr.getElementsByTagName("*")).find(el =>
          el.tagName.endsWith(":defRPr") || el.localName === "defRPr"
        );

        if (defRPr) {
          masterStyles.titleStyle = {
            fontSize: defRPr.getAttribute('sz') || '4400',
            bold: defRPr.getAttribute('b') === '1',
            italic: defRPr.getAttribute('i') === '1',
            color: '',
            typeface: ''
          };

          const solidFill = Array.from(defRPr.getElementsByTagName("*")).find(el =>
            el.tagName.endsWith(":solidFill") || el.localName === "solidFill"
          );
          if (solidFill) {
            const srgbClr = Array.from(solidFill.getElementsByTagName("*")).find(el =>
              el.tagName.endsWith(":srgbClr") || el.localName === "srgbClr"
            );
            if (srgbClr) {
              masterStyles.titleStyle.color = srgbClr.getAttribute('val') || '';
            }
          }

          const latin = Array.from(defRPr.getElementsByTagName("*")).find(el =>
            el.tagName.endsWith(":latin") || el.localName === "latin"
          );
          if (latin) {
            masterStyles.titleStyle.typeface = latin.getAttribute('typeface') || '';
          }
        }
      }
    }

    // 本文スタイルを抽出（レベル1-5）
    const bodyStyle = Array.from(masterDoc.getElementsByTagName("*")).find(el =>
      el.tagName.endsWith(":bodyStyle") || el.localName === "bodyStyle"
    );

    if (bodyStyle) {
      masterStyles.bodyStyle = {};

      for (let level = 1; level <= 5; level++) {
        const lvlPPr = Array.from(bodyStyle.getElementsByTagName("*")).find(el =>
          el.tagName.endsWith(`:lvl${level}pPr`) || el.localName === `lvl${level}pPr`
        );

        if (lvlPPr) {
          const defRPr = Array.from(lvlPPr.getElementsByTagName("*")).find(el =>
            el.tagName.endsWith(":defRPr") || el.localName === "defRPr"
          );

          if (defRPr) {
            masterStyles.bodyStyle[`level${level}`] = {
              fontSize: defRPr.getAttribute('sz') || '1800',
              bold: defRPr.getAttribute('b') === '1',
              italic: defRPr.getAttribute('i') === '1',
              color: '',
              typeface: ''
            };

            const solidFill = Array.from(defRPr.getElementsByTagName("*")).find(el =>
              el.tagName.endsWith(":solidFill") || el.localName === "solidFill"
            );
            if (solidFill) {
              const srgbClr = Array.from(solidFill.getElementsByTagName("*")).find(el =>
                el.tagName.endsWith(":srgbClr") || el.localName === "srgbClr"
              );
              if (srgbClr) {
                masterStyles.bodyStyle[`level${level}`].color = srgbClr.getAttribute('val') || '';
              }
            }

            const latin = Array.from(defRPr.getElementsByTagName("*")).find(el =>
              el.tagName.endsWith(":latin") || el.localName === "latin"
            );
            if (latin) {
              masterStyles.bodyStyle[`level${level}`].typeface = latin.getAttribute('typeface') || '';
            }
          }
        }
      }
    }

    console.log(`マスタースタイル抽出: タイトル=${masterStyles.titleStyle ? 'あり' : 'なし'}, 本文=${Object.keys(masterStyles.bodyStyle || {}).length}レベル`);
    return masterStyles;

  } catch (err) {
    console.log('マスタースタイルの抽出エラー:', err.message);
    return { titleStyle: null, bodyStyle: null };
  }
}

// レイアウトからプレースホルダーの位置を取得
async function getLayoutPosition(zip, slidePath, phType, phIdx) {
  try {
    const slideNum = slidePath.match(/slide(\d+)\.xml/)[1];
    const slideRelsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`;
    const slideRelsFile = zip.file(slideRelsPath);

    if (!slideRelsFile) {
      return null;
    }

    const slideRelsXml = await slideRelsFile.async('string');
    const slideRelsDoc = new DOMParser().parseFromString(slideRelsXml, 'application/xml');

    const layoutRel = Array.from(slideRelsDoc.getElementsByTagName('Relationship')).find(rel =>
      rel.getAttribute('Type').includes('slideLayout')
    );

    if (!layoutRel) {
      return null;
    }

    const layoutPath = `ppt/${layoutRel.getAttribute('Target').replace('../', '')}`;
    const layoutFile = zip.file(layoutPath);

    if (!layoutFile) {
      return null;
    }

    const layoutXml = await layoutFile.async('string');
    const layoutDoc = new DOMParser().parseFromString(layoutXml, 'application/xml');

    const shapes = Array.from(layoutDoc.getElementsByTagNameNS("*", "sp"));

    for (const shape of shapes) {
      const ph = Array.from(shape.getElementsByTagName("*")).find(el =>
        el.tagName.endsWith(":ph") || el.localName === "ph"
      );

      if (!ph) continue;

      const layoutPhType = ph.getAttribute('type');
      const layoutPhIdx = ph.getAttribute('idx');

      const typeMatches = layoutPhType === phType;
      const idxMatches = !phIdx || !layoutPhIdx || layoutPhIdx === phIdx;

      if (typeMatches && idxMatches) {
        const spPr = Array.from(shape.getElementsByTagName("*")).find(el =>
          el.tagName.endsWith(":spPr") || el.localName === "spPr"
        );

        if (!spPr) continue;

        const xfrm = Array.from(spPr.getElementsByTagName("*")).find(el =>
          el.tagName.endsWith(":xfrm") || el.localName === "xfrm"
        );

        if (!xfrm) continue;

        const off = Array.from(xfrm.getElementsByTagName("*")).find(el =>
          el.tagName.endsWith(":off") || el.localName === "off"
        );
        const ext = Array.from(xfrm.getElementsByTagName("*")).find(el =>
          el.tagName.endsWith(":ext") || el.localName === "ext"
        );

        if (off && ext) {
          return {
            position: {
              x: parseInt(off.getAttribute('x') || '0', 10),
              y: parseInt(off.getAttribute('y') || '0', 10)
            },
            size: {
              w: parseInt(ext.getAttribute('cx') || '0', 10),
              h: parseInt(ext.getAttribute('cy') || '0', 10)
            }
          };
        }
      }
    }

    return null;
  } catch (err) {
    console.log('レイアウト位置取得エラー:', err.message);
    return null;
  }
}

// 要素抽出関数（テキストボックス、図形）- 完全版
async function extractElements(doc, themeColors, masterStyles, zip, slidePath) {
  const elements = [];
  const shapes = Array.from(doc.getElementsByTagNameNS("*", "sp"));

  console.log(`${shapes.length}個のshape要素を発見`);

  for (let index = 0; index < shapes.length; index++) {
    const shape = shapes[index];
    try {
      const element = {
        index: index,
        text: '',
        paragraphs: [],
        position: { x: 0, y: 0 },
        size: { width: 0, height: 0 },
        style: {
          fontSize: '',
          color: '',
          bold: false,
          italic: false,
          typeface: '',
          alignment: ''
        },
        fillColor: '',
        borderColor: '',
        borderWidth: 0,
        placeholderType: null
      };

      // プレースホルダーかどうかを確認
      const ph = Array.from(shape.getElementsByTagName("*")).find(el =>
        el.tagName.endsWith(":ph") || el.localName === "ph"
      );

      let phType = null;
      let phIdx = null;

      if (ph) {
        phType = ph.getAttribute('type');
        phIdx = ph.getAttribute('idx');
        element.placeholderType = phType;
        console.log(`要素${index}: プレースホルダータイプ=${phType}, idx=${phIdx}`);
      }

      // テキスト取得（改行を保持）
      const txBody = Array.from(shape.getElementsByTagName("*")).find(el =>
        el.tagName.endsWith(":txBody") || el.localName === "txBody"
      );

      if (txBody) {
        const paragraphs = Array.from(txBody.getElementsByTagName("*")).filter(el =>
          el.tagName.endsWith(":p") || el.localName === "p"
        );

        paragraphs.forEach(p => {
          const paragraphData = {
            text: '',
            level: 0,
            bullet: null
          };

          // 段落プロパティ（<a:pPr>）から箇条書き情報を取得
          const pPr = Array.from(p.childNodes).find(node => {
            if (node.nodeType === 1) {
              const tagName = node.tagName || node.localName;
              return tagName && (tagName.endsWith(':pPr') || tagName === 'pPr');
            }
            return false;
          });

          if (pPr) {
            // インデントレベルを取得
            const marL = pPr.getAttribute('marL');
            if (marL) {
              paragraphData.level = Math.round(parseInt(marL, 10) / 285750);
            }

            // 箇条書きマーカーを確認
            const buChar = Array.from(pPr.getElementsByTagName("*")).find(el =>
              el.tagName.endsWith(":buChar") || el.localName === "buChar"
            );

            const buFont = Array.from(pPr.getElementsByTagName("*")).find(el =>
              el.tagName.endsWith(":buFont") || el.localName === "buFont"
            );

            const buAutoNum = Array.from(pPr.getElementsByTagName("*")).find(el =>
              el.tagName.endsWith(":buAutoNum") || el.localName === "buAutoNum"
            );

            if (buChar || buFont || buAutoNum) {
              paragraphData.bullet = {
                type: buAutoNum ? 'number' : 'char',
                char: buChar ? (buChar.getAttribute('char') || '•') : '•',
                font: buFont ? (buFont.getAttribute('typeface') || '') : '',
                numType: buAutoNum ? (buAutoNum.getAttribute('type') || 'arabicPeriod') : undefined
              };
            }
          }

          // 段落内のテキストを取得
          Array.from(p.childNodes).forEach(node => {
            if (node.nodeType === 1) {
              const tagName = node.tagName || node.localName;
              if (tagName && (tagName.endsWith(':r') || tagName === 'r')) {
                const tNode = Array.from(node.getElementsByTagName("*")).find(el =>
                  el.tagName.endsWith(":t") || el.localName === "t"
                );
                if (tNode) {
                  paragraphData.text += tNode.textContent || '';
                }
              } else if (tagName && (tagName.endsWith(':br') || tagName === 'br')) {
                paragraphData.text += '\n';
              }
            }
          });

          element.paragraphs.push(paragraphData);
        });

        element.text = element.paragraphs.map(p => p.text).join('\n');
      }

      // 位置とサイズ取得
      const xfrm = Array.from(shape.getElementsByTagName("*")).find(el =>
        el.tagName.endsWith(":xfrm") || el.localName === "xfrm"
      );

      let hasPosition = false;
      if (xfrm) {
        const off = Array.from(xfrm.getElementsByTagName("*")).find(el =>
          el.tagName.endsWith(":off") || el.localName === "off"
        );
        const ext = Array.from(xfrm.getElementsByTagName("*")).find(el =>
          el.tagName.endsWith(":ext") || el.localName === "ext"
        );

        if (off) {
          element.position.x = parseInt(off.getAttribute('x') || '0');
          element.position.y = parseInt(off.getAttribute('y') || '0');
          hasPosition = true;
        }
        if (ext) {
          element.size.width = parseInt(ext.getAttribute('cx') || '0');
          element.size.height = parseInt(ext.getAttribute('cy') || '0');
        }
      }

      // プレースホルダーで位置情報がない場合、レイアウトから取得
      if (!hasPosition && phType) {
        const layoutPosition = await getLayoutPosition(zip, slidePath, phType, phIdx);
        if (layoutPosition) {
          element.position.x = layoutPosition.position.x;
          element.position.y = layoutPosition.position.y;
          element.size.width = layoutPosition.size.w;
          element.size.height = layoutPosition.size.h;
          console.log(`要素${index}: レイアウトから位置を取得 x=${element.position.x}, y=${element.position.y}`);
        }
      }

      // スタイル情報取得
      const rPr = Array.from(shape.getElementsByTagName("*")).find(el =>
        el.tagName.endsWith(":rPr") || el.localName === "rPr"
      );

      // プレースホルダーの場合、マスターからデフォルトスタイルを取得
      let masterStyle = null;
      if (phType && masterStyles) {
        if (phType === 'title' || phType === 'ctrTitle') {
          masterStyle = masterStyles.titleStyle;
          console.log(`要素${index}: タイトルプレースホルダー、マスタースタイル適用`);
        } else if (phType === 'body') {
          masterStyle = masterStyles.bodyStyle?.level1;
          console.log(`要素${index}: 本文プレースホルダー、マスタースタイル適用`);
        }
      }

      if (masterStyle) {
        element.style.fontSize = masterStyle.fontSize;
        element.style.bold = masterStyle.bold;
        element.style.italic = masterStyle.italic;
        element.style.color = masterStyle.color;
        element.style.typeface = masterStyle.typeface;
      }

      // スライド固有のスタイルで上書き
      if (rPr) {
        const sz = rPr.getAttribute('sz');
        if (sz) {
          element.style.fontSize = sz;
        }

        const b = rPr.getAttribute('b');
        if (b) {
          element.style.bold = b === '1';
        }

        const i = rPr.getAttribute('i');
        if (i) {
          element.style.italic = i === '1';
        }

        const solidFill = Array.from(rPr.getElementsByTagName("*")).find(el =>
          el.tagName.endsWith(":solidFill") || el.localName === "solidFill"
        );
        if (solidFill) {
          element.style.color = extractColor(solidFill, themeColors);
        }

        const latin = Array.from(rPr.getElementsByTagName("*")).find(el =>
          el.tagName.endsWith(":latin") || el.localName === "latin"
        );
        if (latin) {
          element.style.typeface = latin.getAttribute('typeface') || '';
        }
      }

      // テキスト配置
      const pPr = Array.from(shape.getElementsByTagName("*")).find(el =>
        el.tagName.endsWith(":pPr") || el.localName === "pPr"
      );
      if (pPr) {
        const algn = pPr.getAttribute('algn');
        if (algn) {
          element.style.alignment = algn;
        }
      }

      // 図形プロパティ（背景色・枠線）
      const spPr = Array.from(shape.getElementsByTagName("*")).find(el =>
        el.tagName.endsWith(":spPr") || el.tagName === "spPr" || el.localName === "spPr"
      );

      if (spPr) {
        const noFill = Array.from(spPr.childNodes).find(node => {
          if (node.nodeType === 1) {
            const tagName = node.tagName || node.localName;
            return tagName && (tagName.endsWith(':noFill') || tagName === 'noFill');
          }
          return false;
        });

        const shapeSolidFill = Array.from(spPr.childNodes).find(node => {
          if (node.nodeType === 1) {
            const tagName = node.tagName || node.localName;
            return tagName && (tagName.endsWith(':solidFill') || tagName === 'solidFill');
          }
          return false;
        });

        if (!noFill && shapeSolidFill) {
          element.fillColor = extractColor(shapeSolidFill, themeColors);
        }

        const ln = Array.from(spPr.getElementsByTagName("*")).find(el =>
          el.tagName.endsWith(":ln") || el.localName === "ln"
        );

        if (ln) {
          const w = ln.getAttribute('w');
          element.borderWidth = w ? parseInt(w) : 12700;

          const lnNoFill = Array.from(ln.childNodes).find(node => {
            if (node.nodeType === 1) {
              const tagName = node.tagName || node.localName;
              return tagName && (tagName.endsWith(':noFill') || tagName === 'noFill');
            }
            return false;
          });

          const borderSolidFill = Array.from(ln.getElementsByTagName("*")).find(el =>
            el.tagName.endsWith(":solidFill") || el.localName === "solidFill"
          );

          if (!lnNoFill && borderSolidFill) {
            element.borderColor = extractColor(borderSolidFill, themeColors);
          }
        }
      }

      if (element.text.trim() || element.fillColor || element.borderColor || element.borderWidth > 0) {
        elements.push(element);
      }

    } catch (err) {
      console.log(`要素${index}の処理中にエラー:`, err.message);
    }
  }

  return elements;
}

// 表抽出関数 - 完全版
function extractTables(doc, themeColors) {
  const tables = [];

  const graphicFrames = Array.from(doc.getElementsByTagNameNS("*", "graphicFrame"));

  console.log(`${graphicFrames.length}個のgraphicFrame要素を発見`);

  graphicFrames.forEach((frame, frameIndex) => {
    try {
      const tbl = Array.from(frame.getElementsByTagName("*")).find(el =>
        el.tagName.endsWith(":tbl") || el.localName === "tbl"
      );

      if (!tbl) return;

      const table = {
        index: frameIndex,
        position: { x: 0, y: 0 },
        size: { width: 0, height: 0 },
        rows: [],
        columnWidths: [],
        hasHeaderRow: false
      };

      const tblPr = Array.from(tbl.getElementsByTagName("*")).find(el =>
        el.tagName.endsWith(":tblPr") || el.localName === "tblPr"
      );

      if (tblPr) {
        const firstRow = tblPr.getAttribute('firstRow');
        table.hasHeaderRow = firstRow === '1';
        if (table.hasHeaderRow) {
          console.log(`表${frameIndex}: ヘッダー行が検出されました`);
        }
      }

      const xfrm = Array.from(frame.getElementsByTagName("*")).find(el =>
        el.tagName.endsWith(":xfrm") || el.localName === "xfrm"
      );

      if (xfrm) {
        const off = Array.from(xfrm.getElementsByTagName("*")).find(el =>
          el.tagName.endsWith(":off") || el.localName === "off"
        );
        const ext = Array.from(xfrm.getElementsByTagName("*")).find(el =>
          el.tagName.endsWith(":ext") || el.localName === "ext"
        );

        if (off) {
          table.position.x = parseInt(off.getAttribute('x') || '0');
          table.position.y = parseInt(off.getAttribute('y') || '0');
        }
        if (ext) {
          table.size.width = parseInt(ext.getAttribute('cx') || '0');
          table.size.height = parseInt(ext.getAttribute('cy') || '0');
        }
      }

      const tblGrid = Array.from(tbl.getElementsByTagName("*")).find(el =>
        el.tagName.endsWith(":tblGrid") || el.localName === "tblGrid"
      );

      if (tblGrid) {
        const gridCols = Array.from(tblGrid.getElementsByTagName("*")).filter(el =>
          el.tagName.endsWith(":gridCol") || el.localName === "gridCol"
        );
        table.columnWidths = gridCols.map(col => parseInt(col.getAttribute('w') || '0'));
      }

      const rows = Array.from(tbl.getElementsByTagName("*")).filter(el =>
        el.tagName.endsWith(":tr") || el.localName === "tr"
      );

      rows.forEach((row, rowIndex) => {
        const rowData = {
          height: parseInt(row.getAttribute('h') || '0'),
          cells: [],
          isHeader: table.hasHeaderRow && rowIndex === 0
        };

        const cells = Array.from(row.getElementsByTagName("*")).filter(el =>
          el.tagName.endsWith(":tc") || el.localName === "tc"
        );

        cells.forEach((cell, cellIndex) => {
          const cellData = {
            text: '',
            style: {
              fontSize: '',
              color: '',
              bold: false,
              italic: false,
              typeface: '',
              alignment: 'l'
            },
            fill: {
              color: ''
            },
            borders: {
              left: { width: 0, color: '', dashType: 'solid' },
              right: { width: 0, color: '', dashType: 'solid' },
              top: { width: 0, color: '', dashType: 'solid' },
              bottom: { width: 0, color: '', dashType: 'solid' }
            },
            margins: {
              left: 0,
              right: 0,
              top: 0,
              bottom: 0
            }
          };

          const txBody = Array.from(cell.getElementsByTagName("*")).find(el =>
            el.tagName.endsWith(":txBody") || el.localName === "txBody"
          );

          if (txBody) {
            const paragraphs = Array.from(txBody.getElementsByTagName("*")).filter(el =>
              el.tagName.endsWith(":p") || el.localName === "p"
            );

            const paragraphTexts = paragraphs.map(p => {
              let paragraphText = '';
              Array.from(p.childNodes).forEach(node => {
                if (node.nodeType === 1) {
                  const tagName = node.tagName || node.localName;
                  if (tagName && (tagName.endsWith(':r') || tagName === 'r')) {
                    const tNode = Array.from(node.getElementsByTagName("*")).find(el =>
                      el.tagName.endsWith(":t") || el.localName === "t"
                    );
                    if (tNode) {
                      paragraphText += tNode.textContent || '';
                    }
                  } else if (tagName && (tagName.endsWith(':br') || tagName === 'br')) {
                    paragraphText += '\n';
                  }
                }
              });
              return paragraphText;
            });

            cellData.text = paragraphTexts.join('\n');

            const rPr = Array.from(txBody.getElementsByTagName("*")).find(el =>
              el.tagName.endsWith(":rPr") || el.localName === "rPr"
            );

            if (rPr) {
              const sz = rPr.getAttribute('sz');
              if (sz) cellData.style.fontSize = sz;

              const b = rPr.getAttribute('b');
              if (b) cellData.style.bold = b === '1';

              const i = rPr.getAttribute('i');
              if (i) cellData.style.italic = i === '1';

              const solidFill = Array.from(rPr.getElementsByTagName("*")).find(el =>
                el.tagName.endsWith(":solidFill") || el.localName === "solidFill"
              );

              if (solidFill) {
                cellData.style.color = extractColor(solidFill, themeColors);
              }

              const latin = Array.from(rPr.getElementsByTagName("*")).find(el =>
                el.tagName.endsWith(":latin") || el.localName === "latin"
              );
              if (latin) {
                cellData.style.typeface = latin.getAttribute('typeface') || '';
              }
            }

            const pPr = Array.from(txBody.getElementsByTagName("*")).find(el =>
              el.tagName.endsWith(":pPr") || el.localName === "pPr"
            );
            if (pPr) {
              const algn = pPr.getAttribute('algn');
              if (algn) cellData.style.alignment = algn;
            }
          }

          const tcPr = Array.from(cell.getElementsByTagName("*")).find(el =>
            el.tagName.endsWith(":tcPr") || el.localName === "tcPr"
          );

          if (tcPr) {
            cellData.margins.left = parseInt(tcPr.getAttribute('marL') || '0');
            cellData.margins.right = parseInt(tcPr.getAttribute('marR') || '0');
            cellData.margins.top = parseInt(tcPr.getAttribute('marT') || '0');
            cellData.margins.bottom = parseInt(tcPr.getAttribute('marB') || '0');

            const cellSolidFill = Array.from(tcPr.getElementsByTagName("*")).find(el =>
              (el.tagName.endsWith(":solidFill") || el.localName === "solidFill") &&
              el.parentNode === tcPr
            );

            if (cellSolidFill) {
              cellData.fill.color = extractColor(cellSolidFill, themeColors);
            }

            if (rowData.isHeader && !cellData.fill.color && themeColors) {
              const textColor = cellData.style.color.toUpperCase();
              if (textColor === 'FFFFFF' || textColor === 'FFF' || !textColor) {
                cellData.fill.color = themeColors['tx2'] || themeColors['dk2'] || '44546A';
              } else {
                cellData.fill.color = themeColors['tx1'] || themeColors['dk1'] || '000000';
              }
              console.log(`ヘッダー行のセル${cellIndex}: デフォルト背景色 ${cellData.fill.color} を適用（テキスト色: ${textColor}）`);
            }

            const lnL = Array.from(tcPr.getElementsByTagName("*")).find(el =>
              el.tagName.endsWith(":lnL") || el.localName === "lnL"
            );
            cellData.borders.left = extractBorderInfo(lnL, themeColors);

            const lnR = Array.from(tcPr.getElementsByTagName("*")).find(el =>
              el.tagName.endsWith(":lnR") || el.localName === "lnR"
            );
            cellData.borders.right = extractBorderInfo(lnR, themeColors);

            const lnT = Array.from(tcPr.getElementsByTagName("*")).find(el =>
              el.tagName.endsWith(":lnT") || el.localName === "lnT"
            );
            cellData.borders.top = extractBorderInfo(lnT, themeColors);

            const lnB = Array.from(tcPr.getElementsByTagName("*")).find(el =>
              el.tagName.endsWith(":lnB") || el.localName === "lnB"
            );
            cellData.borders.bottom = extractBorderInfo(lnB, themeColors);
          }

          rowData.cells.push(cellData);
        });

        table.rows.push(rowData);
      });

      tables.push(table);
      console.log(`表${frameIndex}: ${table.rows.length}行 × ${table.columnWidths.length}列`);

    } catch (err) {
      console.log(`表${frameIndex}の処理中にエラー:`, err.message);
    }
  });

  return tables;
}

// 線・コネクタ抽出関数 - 完全版
function extractLines(doc, themeColors) {
  const lines = [];

  const connectors = Array.from(doc.getElementsByTagNameNS("*", "cxnSp"));

  console.log(`${connectors.length}個のcxnSp要素（線・コネクタ）を発見`);

  connectors.forEach((cxn, index) => {
    try {
      const line = {
        index: index,
        position: { x: 0, y: 0 },
        size: { width: 0, height: 0 },
        lineWidth: 0,
        lineColor: "",
        lineDash: "solid",
        arrowStart: "none",
        arrowEnd: "none"
      };

      const xfrm = Array.from(cxn.getElementsByTagName("*")).find(el =>
        el.tagName.endsWith(":xfrm") || el.localName === "xfrm"
      );

      let flipH = false;

      if (xfrm) {
        flipH = xfrm.getAttribute("flipH") === "1";

        const off = Array.from(xfrm.getElementsByTagName("*")).find(el =>
          el.tagName.endsWith(":off") || el.localName === "off"
        );
        if (off) {
          line.position.x = parseInt(off.getAttribute("x") || "0", 10);
          line.position.y = parseInt(off.getAttribute("y") || "0", 10);
        }

        const ext = Array.from(xfrm.getElementsByTagName("*")).find(el =>
          el.tagName.endsWith(":ext") || el.localName === "ext"
        );
        if (ext) {
          line.size.width = parseInt(ext.getAttribute("cx") || "0", 10);
          line.size.height = parseInt(ext.getAttribute("cy") || "0", 10);
        }
      }

      const ln = Array.from(cxn.getElementsByTagName("*")).find(el =>
        el.tagName.endsWith(":ln") || el.localName === "ln"
      );

      if (ln) {
        const width = ln.getAttribute("w");
        if (width) {
          line.lineWidth = parseInt(width, 10);
        } else {
          line.lineWidth = 9525;
        }

        const solidFill = Array.from(ln.getElementsByTagName("*")).find(el =>
          el.tagName.endsWith(":solidFill") || el.localName === "solidFill"
        );
        if (solidFill) {
          line.lineColor = extractColor(solidFill, themeColors);
        }

        const prstDash = Array.from(ln.getElementsByTagName("*")).find(el =>
          el.tagName.endsWith(":prstDash") || el.localName === "prstDash"
        );
        if (prstDash) {
          line.lineDash = prstDash.getAttribute("val") || "solid";
        }

        const headEnd = Array.from(ln.getElementsByTagName("*")).find(el =>
          el.tagName.endsWith(":headEnd") || el.localName === "headEnd"
        );
        if (headEnd) {
          line.arrowStart = headEnd.getAttribute("type") || "none";
        }

        const tailEnd = Array.from(ln.getElementsByTagName("*")).find(el =>
          el.tagName.endsWith(":tailEnd") || el.localName === "tailEnd"
        );
        if (tailEnd) {
          line.arrowEnd = tailEnd.getAttribute("type") || "none";
        }
      }

      if (flipH) {
        const temp = line.arrowStart;
        line.arrowStart = line.arrowEnd;
        line.arrowEnd = temp;
        console.log(`線${index}: flipH=1のため矢印を反転`);
      }

      lines.push(line);
      console.log(`線${index}: 位置(${line.position.x}, ${line.position.y}), 太さ=${line.lineWidth}, スタイル=${line.lineDash}, 矢印開始=${line.arrowStart}, 矢印終了=${line.arrowEnd}`);

    } catch (err) {
      console.log(`線${index}の処理中にエラー:`, err.message);
    }
  });

  return lines;
}

// AI用プロンプト付きJSON生成関数
function generatePromptWithJSON(elements, tables, lines, template, slidePath) {
  const data = {
    slide: slidePath.split('/').pop().replace('.xml', ''),
    template: template,
    elements: elements.map((el, index) => {
      const xEmu = el.position?.x || 0;
      const yEmu = el.position?.y || 0;
      const wEmu = el.size?.width || 0;
      const hEmu = el.size?.height || 0;

      const elementData = {
        id: index + 1,
        text: el.text || "",
        x: parseFloat(emuToInch(xEmu).toFixed(3)),
        y: parseFloat(emuToInch(yEmu).toFixed(3)),
        w: parseFloat(emuToInch(wEmu).toFixed(3)),
        h: parseFloat(emuToInch(hEmu).toFixed(3)),
        fontSize: parseFloat(fontSizeToPoint(parseInt(el.style?.fontSize) || 1800).toFixed(1)),
        color: normalizeColorHex(el.style?.color || "000000"),
        bold: el.style?.bold || false,
        italic: el.style?.italic || false,
        font: el.style?.typeface || "",
        align: el.style?.alignment || "l",
        fill: normalizeColorHex(el.fillColor) || "",
        borderColor: normalizeColorHex(el.borderColor) || "",
        borderWidth: parseFloat(emuToPoint(el.borderWidth || 0).toFixed(2))
      };

      if (el.paragraphs && el.paragraphs.length > 0) {
        elementData.paragraphs = el.paragraphs.map(p => ({
          text: p.text,
          level: p.level,
          bullet: p.bullet
        }));
      }

      return elementData;
    }),
    tables: tables.map((tbl, index) => {
      const xEmu = tbl.position?.x || 0;
      const yEmu = tbl.position?.y || 0;
      const wEmu = tbl.size?.width || 0;
      const hEmu = tbl.size?.height || 0;

      return {
        id: index + 1,
        x: parseFloat(emuToInch(xEmu).toFixed(3)),
        y: parseFloat(emuToInch(yEmu).toFixed(3)),
        w: parseFloat(emuToInch(wEmu).toFixed(3)),
        h: parseFloat(emuToInch(hEmu).toFixed(3)),
        hasHeader: tbl.hasHeaderRow || false,
        colW: tbl.columnWidths.map(w => parseFloat(emuToInch(w).toFixed(3))),
        rows: tbl.rows.map(row => ({
          h: parseFloat(emuToInch(row.height).toFixed(3)),
          isHeader: row.isHeader || false,
          cells: row.cells.map(cell => ({
            text: cell.text,
            fontSize: parseFloat(fontSizeToPoint(parseInt(cell.style?.fontSize) || 1400).toFixed(1)),
            color: cell.style?.color || "",
            bold: cell.style?.bold || false,
            italic: cell.style?.italic || false,
            font: cell.style?.typeface || "",
            align: cell.style?.alignment || "l",
            fill: cell.fill?.color || "",
            border: [
              parseFloat(emuToPoint(cell.borders?.top?.width || 0).toFixed(2)),
              parseFloat(emuToPoint(cell.borders?.right?.width || 0).toFixed(2)),
              parseFloat(emuToPoint(cell.borders?.bottom?.width || 0).toFixed(2)),
              parseFloat(emuToPoint(cell.borders?.left?.width || 0).toFixed(2))
            ],
            borderColor: [
              cell.borders?.top?.color || "",
              cell.borders?.right?.color || "",
              cell.borders?.bottom?.color || "",
              cell.borders?.left?.color || "",
            ],
            borderStyle: [
              cell.borders?.top?.dashType || "solid",
              cell.borders?.right?.dashType || "solid",
              cell.borders?.bottom?.dashType || "solid",
              cell.borders?.left?.dashType || "solid"
            ],
            margin: [
              parseFloat(emuToInch(cell.margins?.top || 0).toFixed(3)),
              parseFloat(emuToInch(cell.margins?.right || 0).toFixed(3)),
              parseFloat(emuToInch(cell.margins?.bottom || 0).toFixed(3)),
              parseFloat(emuToInch(cell.margins?.left || 0).toFixed(3))
            ]
          }))
        }))
      };
    }),
    lines: lines.map((line, index) => {
      const xEmu = line.position?.x || 0;
      const yEmu = line.position?.y || 0;
      const wEmu = line.size?.width || 0;
      const hEmu = line.size?.height || 0;

      return {
        id: index + 1,
        x: parseFloat(emuToInch(xEmu).toFixed(3)),
        y: parseFloat(emuToInch(yEmu).toFixed(3)),
        w: parseFloat(emuToInch(wEmu).toFixed(3)),
        h: parseFloat(emuToInch(hEmu).toFixed(3)),
        lineWidth: parseFloat(emuToPoint(line.lineWidth || 0).toFixed(2)),
        lineColor: normalizeColorHex(line.lineColor || "000000"),
        lineDash: line.lineDash || "solid",
        arrowStart: line.arrowStart || "none",
        arrowEnd: line.arrowEnd || "none"
      };
    })
  };

  // AI用の詳細プロンプトを生成
  const prompt = `# PowerPoint Slide Reproduction Task

PptxGenJSで下記のパワポを完全再現して。**位置・サイズ・色・フォント・罫線**全て完璧に。

## JSON Structure

\`\`\`json
{
  "slide": "slide name",
  "template": {template info},  // テンプレート情報
  "elements": [{shape data}],   // 図形・テキストボックス
  "tables": [{table data}],     // 表
  "lines": [{line data}]        // 線・コネクタ
}
\`\`\`

### Template (テンプレート情報)

スライドテンプレートから抽出された情報:
- **background**: 背景色(RGB hex, ""=なし)
- **slideNumber**: スライド番号設定(nullまたはオブジェクト)
  - **x, y, w, h**: 位置とサイズ(インチ)
  - **fontSize**: フォントサイズ(pt)
  - **font**: フォント名
  - **color**: 色(RGB hex)
  - **bold**: 太字
  - **align**: 配置("l"/"ctr"/"r")
- **fixedImages**: 固定画像配列(ロゴなど)
  - **id, name**: ID・名前
  - **x, y, w, h**: 位置とサイズ(インチ)

### Elements (図形・テキストボックス)

各要素:
- **id**: ID
- **text**: テキスト内容（全段落を結合）
- **x, y**: 位置(インチ)
- **w, h**: サイズ(インチ)
- **fontSize**: フォントサイズ(pt)
- **color**: テキスト色(RGB hex, "000000"=黒)
- **bold, italic**: スタイル
- **font**: フォント名
- **align**: 配置("l"=左, "ctr"=中央, "r"=右)
- **fill**: 背景色(RGB hex, ""=透明)
- **borderColor**: 枠線色(RGB hex, ""=枠線なし)
- **borderWidth**: 枠線太さ(pt)
- **paragraphs**: 段落配列（箇条書き含む、存在する場合のみ）
  - **text**: 段落テキスト
  - **level**: インデントレベル(0=なし, 1以上=箇条書きレベル)
  - **bullet**: 箇条書き情報(nullは箇条書きなし)
    - **type**: "char"=文字マーカー, "number"=番号付き
    - **char**: 箇条書き文字（例: "•", "n"）
    - **font**: 箇条書きフォント（例: "Wingdings"）
    - **numType**: 番号タイプ（type="number"の場合、例: "arabicPeriod"）

### Tables (表)

各表:
- **id**: ID
- **x, y**: 位置(インチ)
- **w, h**: サイズ(インチ)
- **hasHeader**: ヘッダー行の有無
- **colW**: 列幅配列(インチ)
- **rows**: 行配列
  - **h**: 行高さ(インチ)
  - **isHeader**: ヘッダー行か
  - **cells**: セル配列
    - **text**: テキスト
    - **fontSize**: フォントサイズ(pt)
    - **color**: テキスト色(RGB hex)
    - **bold, italic**: スタイル
    - **font**: フォント名
    - **align**: 配置
    - **fill**: セル背景色(RGB hex)
    - **border**: 罫線太さ配列[top, right, bottom, left](pt)
    - **borderColor**: 罫線色配列[top, right, bottom, left](RGB hex)
    - **borderStyle**: 罫線スタイル配列[top, right, bottom, left]("solid"/"dash"/"dot")
    - **margin**: マージン配列[top, right, bottom, left](インチ)

### Lines (線・コネクタ)

各線:
- **id**: ID
- **x, y**: 開始位置(インチ)
- **w, h**: 幅・高さ(インチ) ※wが線の長さ(水平), hが高さ(垂直)
- **lineWidth**: 線の太さ(pt)
- **lineColor**: 線の色(RGB hex)
- **lineDash**: 線のスタイル("solid"=実線, "dash"=破線, "dot"=点線, "dashDot"=一点鎖線, "lgDash"=長い破線, "sysDot"=システム点線など)
- **arrowStart**: 開始側の矢印("none"=なし, "arrow"=矢印, "triangle"=三角, "diamond"=菱形, "oval"=丸など)
- **arrowEnd**: 終了側の矢印(同上)

---

## PptxGenJS Implementation

### テンプレート設定

\`\`\`javascript
const pptx = new PptxGenJS();
const slide = pptx.addSlide();

// 背景色を設定
if (template.background) {
  slide.background = { color: template.background };
}

// スライド番号を設定
if (template.slideNumber) {
  slide.slideNumber = {
    x: template.slideNumber.x,
    y: template.slideNumber.y,
    fontFace: template.slideNumber.font,
    fontSize: template.slideNumber.fontSize,
    color: template.slideNumber.color,
    bold: template.slideNumber.bold
  };
}

// 固定画像を追加（ロゴなど）
// 注意: 画像データは別途用意する必要があります
template.fixedImages.forEach(img => {
  // slide.addImage({
  //   data: "image/png;base64,..." または path: "logo.png",
  //   x: img.x, y: img.y, w: img.w, h: img.h
  // });
});
\`\`\`

### 図形

\`\`\`javascript
elements.forEach(el => {
  // 箇条書きがある場合は段落ごとに処理
  if (el.paragraphs && el.paragraphs.length > 0) {
    const textContent = el.paragraphs.map(p => {
      const options = {
        bullet: false,
        indentLevel: p.level
      };

      // 箇条書き設定
      if (p.bullet) {
        if (p.bullet.type === 'char') {
          options.bullet = {
            type: p.bullet.char,
            characterCode: p.bullet.char.charCodeAt(0).toString(16)
          };
          if (p.bullet.font) {
            options.bullet.fontFace = p.bullet.font;
          }
        } else if (p.bullet.type === 'number') {
          options.bullet = { type: p.bullet.numType || 'number' };
        }
      }

      return { text: p.text, options };
    });

    slide.addText(textContent, {
      x: el.x, y: el.y, w: el.w, h: el.h,
      fontSize: el.fontSize,
      color: el.color,
      bold: el.bold,
      italic: el.italic,
      fontFace: el.font,
      align: el.align === "l" ? "left" : el.align === "ctr" ? "center" : "right",
      fill: el.fill ? { color: el.fill } : undefined,
      line: el.borderColor ? { color: el.borderColor, pt: el.borderWidth } : undefined
    });
  } else {
    // 箇条書きなしの場合は従来通り
    slide.addText(el.text, {
      x: el.x, y: el.y, w: el.w, h: el.h,
      fontSize: el.fontSize,
      color: el.color,
      bold: el.bold,
      italic: el.italic,
      fontFace: el.font,
      align: el.align === "l" ? "left" : el.align === "ctr" ? "center" : "right",
      fill: el.fill ? { color: el.fill } : undefined,
      line: el.borderColor ? { color: el.borderColor, pt: el.borderWidth } : undefined,
      breakLine: true
    });
  }
});
\`\`\`

### 表

\`\`\`javascript
tables.forEach(table => {
  const tableData = table.rows.map(row =>
    row.cells.map(cell => ({
      text: cell.text,
      options: {
        fontSize: cell.fontSize,
        color: cell.color,
        bold: cell.bold,
        italic: cell.italic,
        fontFace: cell.font,
        align: cell.align === "l" ? "left" : cell.align === "ctr" ? "center" : "right",
        fill: cell.fill ? { color: cell.fill } : undefined,
        border: [
          { pt: cell.border[0], color: cell.borderColor[0], type: cell.borderStyle[0] },
          { pt: cell.border[1], color: cell.borderColor[1], type: cell.borderStyle[1] },
          { pt: cell.border[2], color: cell.borderColor[2], type: cell.borderStyle[2] },
          { pt: cell.border[3], color: cell.borderColor[3], type: cell.borderStyle[3] }
        ],
        margin: cell.margin
      }
    }))
  );

  slide.addTable(tableData, {
    x: table.x, y: table.y, w: table.w,
    colW: table.colW,
    rowH: table.rows.map(r => r.h)
  });
});
\`\`\`

### 線・コネクタ

\`\`\`javascript
lines.forEach(line => {
  slide.addShape("line", {
    x: line.x,
    y: line.y,
    w: line.w,
    h: line.h,
    line: {
      color: line.lineColor,
      pt: line.lineWidth,
      dashType: line.lineDash,
      beginArrowType: line.arrowStart === "none" ? undefined : line.arrowStart,
      endArrowType: line.arrowEnd === "none" ? undefined : line.arrowEnd
    }
  });
});
\`\`\`

---

## Important Notes

1. **Colors**: 6-digit RGB hex ("FF0000"=red, ""=none)
2. **Alignment**: "l"→"left", "ctr"→"center", "r"→"right"
3. **Units**: All positions/sizes in inches, fonts in points
4. **Empty values**: ""=transparent/no border, 0=no border
5. **Header rows**: hasHeader=true means row 1 is header (already has proper background color)
6. **Line styles**: "solid", "dash", "dot", "dashDot", "lgDash", "lgDashDot", "sysDash", "sysDot"
7. **Arrow types**: "none", "arrow", "triangle", "diamond", "oval", "stealth"
8. **Line breaks**: Text contains "\\n" for line breaks. Use breakLine: true in PptxGenJS addText options
9. **Bullets**: Use paragraphs array for bullet points. level indicates indent (0=none, 1+=levels). bullet.char with bullet.font (e.g., Wingdings) for custom markers

---

## Slide Data (JSON)

\`\`\`json
${JSON.stringify(data, null, 2)}
\`\`\`

---

**完璧に再現してください。位置・サイズ・色・フォントが1pxでもずれないように。**`;

  return prompt;
}

// メイン解析関数
export async function analyzePPTX(file) {
  try {
    // JSZipはcontent_scriptとして既にロード済み
    if (typeof JSZip === 'undefined') {
      throw new Error('JSZip is not loaded. Please check manifest.json content_scripts configuration.');
    }

    const buf = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(buf);

    // テーマカラーを読み込む
    const themeColors = await loadThemeColors(zip);

    // スライド一覧を取得
    const slideFiles = Object.keys(zip.files)
      .filter(p => /^ppt\/slides\/slide\d+\.xml$/i.test(p))
      .sort((a, b) => {
        const na = parseInt(a.match(/slide(\d+)\.xml/i)[1], 10);
        const nb = parseInt(b.match(/slide(\d+)\.xml/i)[1], 10);
        return na - nb;
      });

    if (slideFiles.length === 0) {
      throw new Error("スライドが見つかりませんでした");
    }

    const results = [];

    for (const slidePath of slideFiles) {
      const xmlStr = await zip.file(slidePath).async("string");
      const parser = new DOMParser();
      const doc = parser.parseFromString(xmlStr, "application/xml");

      const template = await extractTemplateInfo(zip, slidePath, themeColors);
      const masterStyles = await extractMasterStyles(zip, slidePath);
      const elements = await extractElements(doc, themeColors, masterStyles, zip, slidePath);
      const tables = extractTables(doc, themeColors);
      const lines = extractLines(doc, themeColors);

      const promptWithJson = generatePromptWithJSON(elements, tables, lines, template, slidePath);

      results.push({
        slideNumber: parseInt(slidePath.match(/slide(\d+)\.xml/i)[1], 10),
        slidePath: slidePath,
        promptWithJson: promptWithJson,
        elementCount: elements.length,
        tableCount: tables.length,
        lineCount: lines.length
      });
    }

    return {
      success: true,
      fileName: file.name,
      slideCount: slideFiles.length,
      slides: results
    };

  } catch (err) {
    return {
      success: false,
      error: err.message
    };
  }
}
