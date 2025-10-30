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

// グループ変換情報を取得する関数
function getGroupTransform(grpSpElement) {
  const grpSpPr = Array.from(grpSpElement.getElementsByTagName("*")).find(el =>
    el.tagName.endsWith(":grpSpPr") || el.localName === "grpSpPr"
  );

  if (!grpSpPr) return null;

  const xfrm = Array.from(grpSpPr.getElementsByTagName("*")).find(el =>
    el.tagName.endsWith(":xfrm") || el.localName === "xfrm"
  );

  if (!xfrm) return null;

  const transform = {
    off_x: 0,
    off_y: 0,
    ext_cx: 1,
    ext_cy: 1,
    chOff_x: 0,
    chOff_y: 0,
    chExt_cx: 1,
    chExt_cy: 1
  };

  const off = Array.from(xfrm.getElementsByTagName("*")).find(el =>
    el.tagName.endsWith(":off") || el.localName === "off"
  );
  const ext = Array.from(xfrm.getElementsByTagName("*")).find(el =>
    el.tagName.endsWith(":ext") || el.localName === "ext"
  );
  const chOff = Array.from(xfrm.getElementsByTagName("*")).find(el =>
    el.tagName.endsWith(":chOff") || el.localName === "chOff"
  );
  const chExt = Array.from(xfrm.getElementsByTagName("*")).find(el =>
    el.tagName.endsWith(":chExt") || el.localName === "chExt"
  );

  if (off) {
    transform.off_x = parseInt(off.getAttribute('x') || '0');
    transform.off_y = parseInt(off.getAttribute('y') || '0');
  }
  if (ext) {
    transform.ext_cx = parseInt(ext.getAttribute('cx') || '1');
    transform.ext_cy = parseInt(ext.getAttribute('cy') || '1');
  }
  if (chOff) {
    transform.chOff_x = parseInt(chOff.getAttribute('x') || '0');
    transform.chOff_y = parseInt(chOff.getAttribute('y') || '0');
  }
  if (chExt) {
    transform.chExt_cx = parseInt(chExt.getAttribute('cx') || '1');
    transform.chExt_cy = parseInt(chExt.getAttribute('cy') || '1');
  }

  return transform;
}

// グループの子座標系から絶対座標への変換関数
function transformCoordinates(localPosition, localSize, groupTransform) {
  if (!groupTransform) {
    return { position: localPosition, size: localSize };
  }

  // スケール係数を計算（ゼロ除算を防ぐ）
  const scaleX = groupTransform.chExt_cx !== 0
    ? groupTransform.ext_cx / groupTransform.chExt_cx
    : 1;
  const scaleY = groupTransform.chExt_cy !== 0
    ? groupTransform.ext_cy / groupTransform.chExt_cy
    : 1;

  // 絶対座標を計算
  const absX = groupTransform.off_x + (localPosition.x - groupTransform.chOff_x) * scaleX;
  const absY = groupTransform.off_y + (localPosition.y - groupTransform.chOff_y) * scaleY;
  const absCx = localSize.width * scaleX;
  const absCy = localSize.height * scaleY;

  return {
    position: { x: absX, y: absY },
    size: { width: absCx, height: absCy }
  };
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

// 単一図形の抽出（グループ内外で再利用可能）
async function extractSingleShape(shape, index, themeColors, masterStyles, zip, slidePath) {
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
        placeholderType: null,
        shapeType: 'rect'  // デフォルトは四角形
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
            bullet: null,
            runs: []  // 追加: 各テキストランの配列
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

          // 段落内の各テキストランを個別に解析
          Array.from(p.childNodes).forEach(node => {
            if (node.nodeType === 1) {
              const tagName = node.tagName || node.localName;

              // テキストラン <a:r>
              if (tagName && (tagName.endsWith(':r') || tagName === 'r')) {
                const runData = {
                  text: '',
                  fontSize: 1800,  // デフォルト18pt
                  color: '000000',
                  bold: false,
                  italic: false,
                  font: '',
                  underline: 'none',
                  baseline: 0  // 上付き・下付き文字用
                };

                // ランプロパティ <a:rPr> を取得
                const rPr = Array.from(node.getElementsByTagName("*")).find(el =>
                  el.tagName.endsWith(":rPr") || el.localName === "rPr"
                );

                if (rPr) {
                  // フォントサイズ
                  const sz = rPr.getAttribute('sz');
                  if (sz) runData.fontSize = parseInt(sz);

                  // 太字
                  const b = rPr.getAttribute('b');
                  if (b === '1') runData.bold = true;

                  // 斜体
                  const i = rPr.getAttribute('i');
                  if (i === '1') runData.italic = true;

                  // 下線
                  const u = rPr.getAttribute('u');
                  if (u) runData.underline = u;

                  // 上付き・下付き文字
                  const baseline = rPr.getAttribute('baseline');
                  if (baseline) runData.baseline = parseInt(baseline);

                  // 色
                  const solidFill = Array.from(rPr.getElementsByTagName("*")).find(el =>
                    el.tagName.endsWith(":solidFill") || el.localName === "solidFill"
                  );
                  if (solidFill) {
                    runData.color = extractColor(solidFill, themeColors);
                  }

                  // フォント
                  const latin = Array.from(rPr.getElementsByTagName("*")).find(el =>
                    el.tagName.endsWith(":latin") || el.localName === "latin"
                  );
                  if (latin) {
                    runData.font = latin.getAttribute('typeface') || '';
                  }
                }

                // テキスト取得 <a:t>
                const tNode = Array.from(node.getElementsByTagName("*")).find(el =>
                  el.tagName.endsWith(":t") || el.localName === "t"
                );
                if (tNode) {
                  runData.text = tNode.textContent || '';
                  paragraphData.text += runData.text;  // 全体のテキストにも追加
                }

                // ランをリストに追加（空でない場合のみ）
                if (runData.text) {
                  paragraphData.runs.push(runData);
                }
              }
              // 改行 <a:br>
              else if (tagName && (tagName.endsWith(':br') || tagName === 'br')) {
                paragraphData.text += '\n';
                paragraphData.runs.push({
                  text: '\n',
                  isBreak: true
                });
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

      // 図形プロパティ（背景色・枠線・形状タイプ）
      const spPr = Array.from(shape.getElementsByTagName("*")).find(el =>
        el.tagName.endsWith(":spPr") || el.tagName === "spPr" || el.localName === "spPr"
      );

      if (spPr) {
        // 形状タイプを取得（prstGeom）
        const prstGeom = Array.from(spPr.getElementsByTagName("*")).find(el =>
          el.tagName.endsWith(":prstGeom") || el.localName === "prstGeom"
        );

        if (prstGeom) {
          const prst = prstGeom.getAttribute('prst');
          if (prst) {
            element.shapeType = prst;  // "ellipse", "rect", "roundRect", etc.
          }
        }

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
          // <a:noFill>が存在する場合は枠線なし
          const lnNoFill = Array.from(ln.childNodes).find(node => {
            if (node.nodeType === 1) {
              const tagName = node.tagName || node.localName;
              return tagName && (tagName.endsWith(':noFill') || tagName === 'noFill');
            }
            return false;
          });

          if (lnNoFill) {
            // 枠線なし
            element.borderWidth = 0;
            element.borderColor = '';
          } else {
            // 枠線あり
            const w = ln.getAttribute('w');
            element.borderWidth = w ? parseInt(w) : 12700;

            const borderSolidFill = Array.from(ln.getElementsByTagName("*")).find(el =>
              el.tagName.endsWith(":solidFill") || el.localName === "solidFill"
            );

            if (borderSolidFill) {
              element.borderColor = extractColor(borderSolidFill, themeColors);
            }
          }
        }

        // スタイル参照による枠線のチェック（<p:style><a:lnRef>）
        // 枠線がまだ検出されていない場合、スタイル参照を確認
        if (element.borderWidth === 0 || !element.borderColor) {
          const style = Array.from(shape.getElementsByTagName("*")).find(el =>
            el.tagName.endsWith(":style") || el.localName === "style"
          );

          if (style) {
            const lnRef = Array.from(style.getElementsByTagName("*")).find(el =>
              el.tagName.endsWith(":lnRef") || el.localName === "lnRef"
            );

            if (lnRef) {
              // スタイル参照がある = 枠線あり
              if (element.borderWidth === 0) {
                element.borderWidth = 12700; // デフォルト枠線幅を設定
              }

              // 色情報の抽出を試みる
              if (!element.borderColor) {
                const schemeClr = Array.from(lnRef.getElementsByTagName("*")).find(el =>
                  el.tagName.endsWith(":schemeClr") || el.localName === "schemeClr"
                );
                if (schemeClr) {
                  element.borderColor = extractColor(lnRef, themeColors);
                }
              }
            }
          }
        }
      }

      // 以下のいずれかの条件を満たす図形を返す：
      // 1. テキストがある
      // 2. 背景色がある
      // 3. 枠線がある（色または幅が設定されている）
      // 4. デフォルト以外の形状タイプ（ellipseなど）
      const hasContent = element.text.trim().length > 0;
      const hasFill = element.fillColor && element.fillColor.length > 0;
      const hasBorder = (element.borderColor && element.borderColor.length > 0) || element.borderWidth > 0;
      const hasNonRectShape = element.shapeType && element.shapeType !== 'rect';

      if (hasContent || hasFill || hasBorder || hasNonRectShape) {
        return element;
      }

      return null;

    } catch (err) {
      console.log(`要素${index}の処理中にエラー:`, err.message);
      return null;
    }
}

// 要素抽出関数（テキストボックス、図形）- グループ外の図形のみ
async function extractElements(doc, themeColors, masterStyles, zip, slidePath) {
  const elements = [];
  const shapes = Array.from(doc.getElementsByTagNameNS("*", "sp"));

  console.log(`${shapes.length}個のshape要素を発見`);

  for (let index = 0; index < shapes.length; index++) {
    const shape = shapes[index];

    // グループ内の図形は除外（親要素がgrpSpかチェック）
    const parent = shape.parentElement;
    const parentTag = parent ? (parent.tagName || parent.localName) : null;
    if (parentTag && (parentTag.endsWith(':grpSp') || parentTag === 'grpSp')) {
      continue;  // グループ内の図形はスキップ
    }

    const element = await extractSingleShape(shape, index, themeColors, masterStyles, zip, slidePath);
    if (element) {
      elements.push(element);
    }
  }

  return elements;
}

// 単一表の抽出（グループ内外で再利用可能）
function extractSingleTable(frame, frameIndex, themeColors) {
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
          // セル結合属性を取得（<a:tc>要素に直接付与されている）
          const gridSpan = cell.getAttribute('gridSpan');
          const rowSpan = cell.getAttribute('rowSpan');
          const hMerge = cell.getAttribute('hMerge');
          const vMerge = cell.getAttribute('vMerge');

          // 継続セル判定 - プレースホルダーとして保持（位置ずれを防ぐため）
          const isMergedContinuation = (hMerge === '1' || vMerge === '1');
          if (isMergedContinuation) {
            console.log(`表${frameIndex} 行${rowIndex} XMLセル${cellIndex}: 結合継続セル - プレースホルダーとして保持`);
          }

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
            },
            // セル結合情報
            colspan: gridSpan ? parseInt(gridSpan) : 1,
            rowspan: rowSpan ? parseInt(rowSpan) : 1,
            // 継続セルフラグ（グリッド位置維持のため重要）
            isMergedContinuation: isMergedContinuation
          };

          // 継続セルはテキストを持たない（プレースホルダーのため）
          if (!isMergedContinuation) {
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

          // 継続セルもプレースホルダーとして保持（グリッド位置維持のため）
          // isMergedContinuationフラグで識別可能
          rowData.cells.push(cellData);
        });

        table.rows.push(rowData);
      });

      console.log(`表${frameIndex}: ${table.rows.length}行 × ${table.columnWidths.length}列（論理セル数: ${table.rows.reduce((sum, row) => sum + row.cells.length, 0)}セル）`);
      return table;

    } catch (err) {
      console.log(`表${frameIndex}の処理中にエラー:`, err.message);
      return null;
    }
}

// 表抽出関数 - グループ外の表のみ
function extractTables(doc, themeColors) {
  const tables = [];
  const graphicFrames = Array.from(doc.getElementsByTagNameNS("*", "graphicFrame"));

  console.log(`${graphicFrames.length}個のgraphicFrame要素を発見`);

  graphicFrames.forEach((frame, frameIndex) => {
    // グループ内の表は除外
    const parent = frame.parentElement;
    const parentTag = parent ? (parent.tagName || parent.localName) : null;
    if (parentTag && (parentTag.endsWith(':grpSp') || parentTag === 'grpSp')) {
      return;  // グループ内の表はスキップ
    }

    const table = extractSingleTable(frame, frameIndex, themeColors);
    if (table) {
      tables.push(table);
    }
  });

  return tables;
}

// 単一線の抽出（グループ内外で再利用可能）
function extractSingleLine(cxn, index, themeColors) {
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

      console.log(`線${index}: 位置(${line.position.x}, ${line.position.y}), 太さ=${line.lineWidth}, スタイル=${line.lineDash}, 矢印開始=${line.arrowStart}, 矢印終了=${line.arrowEnd}`);
      return line;

    } catch (err) {
      console.log(`線${index}の処理中にエラー:`, err.message);
      return null;
    }
}

// 線・コネクタ抽出関数 - グループ外の線のみ
function extractLines(doc, themeColors) {
  const lines = [];
  const connectors = Array.from(doc.getElementsByTagNameNS("*", "cxnSp"));

  console.log(`${connectors.length}個のcxnSp要素（線・コネクタ）を発見`);

  connectors.forEach((cxn, index) => {
    // グループ内の線は除外
    const parent = cxn.parentElement;
    const parentTag = parent ? (parent.tagName || parent.localName) : null;
    if (parentTag && (parentTag.endsWith(':grpSp') || parentTag === 'grpSp')) {
      return;  // グループ内の線はスキップ
    }

    const line = extractSingleLine(cxn, index, themeColors);
    if (line) {
      lines.push(line);
    }
  });

  return lines;
}

// グループ化図形の再帰的抽出
async function extractGroupRecursive(grpSpElement, themeColors, masterStyles, zip, slidePath, groupTransform = null) {
  const groupChildren = {
    elements: [],
    tables: [],
    lines: []
  };

  try {
    // グループの変換情報を取得
    const localGroupTransform = getGroupTransform(grpSpElement);

    // 親グループの変換がある場合は累積変換を適用（ネストグループ対応）
    let effectiveTransform = localGroupTransform;
    if (groupTransform && localGroupTransform) {
      // 累積変換の計算（簡略化: 親→子の順で適用）
      // 完全な実装では行列演算が必要だが、基本的なケースでは順次適用で十分
      effectiveTransform = localGroupTransform;
    }

    // グループ内の直接の子要素を走査
    const children = Array.from(grpSpElement.children);
    let elementIndex = 0;
    let tableIndex = 0;
    let lineIndex = 0;

    for (const child of children) {
      const tagName = child.tagName || child.localName;

      if (tagName && (tagName.endsWith(':sp') || tagName === 'sp')) {
        // 通常の図形
        const element = await extractSingleShape(child, elementIndex++, themeColors, masterStyles, zip, slidePath);
        if (element && effectiveTransform) {
          // 座標変換を適用
          const transformed = transformCoordinates(element.position, element.size, effectiveTransform);
          element.position = transformed.position;
          element.size = transformed.size;
          groupChildren.elements.push(element);
        } else if (element) {
          groupChildren.elements.push(element);
        }

      } else if (tagName && (tagName.endsWith(':graphicFrame') || tagName === 'graphicFrame')) {
        // 表
        const table = extractSingleTable(child, tableIndex++, themeColors);
        if (table && effectiveTransform) {
          // 座標変換を適用
          const transformed = transformCoordinates(table.position, table.size, effectiveTransform);
          table.position = transformed.position;
          table.size = transformed.size;
          groupChildren.tables.push(table);
        } else if (table) {
          groupChildren.tables.push(table);
        }

      } else if (tagName && (tagName.endsWith(':cxnSp') || tagName === 'cxnSp')) {
        // 線・コネクタ
        const line = extractSingleLine(child, lineIndex++, themeColors);
        if (line && effectiveTransform) {
          // 座標変換を適用
          const transformed = transformCoordinates(line.position, line.size, effectiveTransform);
          line.position = transformed.position;
          line.size = transformed.size;
          groupChildren.lines.push(line);
        } else if (line) {
          groupChildren.lines.push(line);
        }

      } else if (tagName && (tagName.endsWith(':grpSp') || tagName === 'grpSp')) {
        // ネストされたグループ - 再帰呼び出し
        const nestedGroup = await extractGroupRecursive(child, themeColors, masterStyles, zip, slidePath, effectiveTransform);
        // ネストされたグループの要素を統合
        groupChildren.elements.push(...nestedGroup.elements);
        groupChildren.tables.push(...nestedGroup.tables);
        groupChildren.lines.push(...nestedGroup.lines);
      }
    }

  } catch (err) {
    console.log('グループ抽出中にエラー:', err.message);
  }

  return groupChildren;
}

// グループ化図形を抽出（トップレベルのグループのみ）
async function extractGroups(doc, themeColors, masterStyles, zip, slidePath) {
  const groupElements = {
    elements: [],
    tables: [],
    lines: []
  };

  try {
    // トップレベルのグループ要素のみ取得（親がgrpSpでないもの）
    const allGroups = Array.from(doc.getElementsByTagNameNS("*", "grpSp"));
    const topLevelGroups = allGroups.filter(grp => {
      const parent = grp.parentElement;
      const parentTag = parent ? (parent.tagName || parent.localName) : null;
      return !parentTag || !(parentTag.endsWith(':grpSp') || parentTag === 'grpSp');
    });

    console.log(`${topLevelGroups.length}個のトップレベルグループ要素を発見`);

    for (const group of topLevelGroups) {
      const groupChildren = await extractGroupRecursive(group, themeColors, masterStyles, zip, slidePath);
      groupElements.elements.push(...groupChildren.elements);
      groupElements.tables.push(...groupChildren.tables);
      groupElements.lines.push(...groupChildren.lines);
    }

    console.log(`グループから抽出: 図形=${groupElements.elements.length}, 表=${groupElements.tables.length}, 線=${groupElements.lines.length}`);

  } catch (err) {
    console.log('グループ抽出中にエラー:', err.message);
  }

  return groupElements;
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

      // alignmentを変換（PptxGenJS形式に）
      const alignmentMap = { "l": "left", "ctr": "center", "r": "right" };
      const align = alignmentMap[el.style?.alignment] || el.style?.alignment || "left";

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
        fontFace: el.style?.typeface || "",  // font → fontFace
        align: align,  // 変換済み
        shapeType: el.shapeType || "rect"  // 形状タイプを追加
      };

      // fillをオブジェクト形式に
      if (el.fillColor) {
        elementData.fill = { color: normalizeColorHex(el.fillColor) };
      } else {
        // fillColorがない場合でも、枠線のみの図形には透明な塗りつぶしを設定
        if (el.borderColor && el.borderWidth > 0) {
          elementData.fill = { color: "FFFFFF", transparency: 99 };
        }
      }

      // lineをオブジェクト形式に（枠線がある場合のみ）
      if (el.borderColor && el.borderWidth > 0) {
        elementData.line = {
          color: normalizeColorHex(el.borderColor),
          pt: parseFloat(emuToPoint(el.borderWidth).toFixed(2))
        };
      }

      if (el.paragraphs && el.paragraphs.length > 0) {
        elementData.paragraphs = el.paragraphs.map(p => ({
          text: p.text,
          level: p.level,
          bullet: p.bullet,
          runs: p.runs ? p.runs.map(run => ({
            text: run.text,
            fontSize: parseFloat(fontSizeToPoint(parseInt(run.fontSize || 1800)).toFixed(1)),
            color: normalizeColorHex(run.color || '000000'),
            bold: run.bold || false,
            italic: run.italic || false,
            fontFace: run.font || '',  // font → fontFace
            underline: run.underline || 'none',
            baseline: run.baseline || 0,
            isBreak: run.isBreak || false
          })) : undefined
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
          // 継続セルも含める（グリッド構造を維持するため、rowHの要素数と一致させる）
          cells: row.cells.map(cell => {
            // alignとvalignを変換
            const alignmentMap = { "l": "left", "ctr": "center", "r": "right" };
            const valignMap = { "t": "top", "m": "middle", "b": "bottom" };
            const align = alignmentMap[cell.style?.alignment] || cell.style?.alignment || "left";
            const valign = valignMap[cell.style?.valign] || cell.style?.valign || "top";

            const cellData = {
              text: cell.text,
              fontSize: parseFloat(fontSizeToPoint(parseInt(cell.style?.fontSize) || 1400).toFixed(1)),
              color: cell.style?.color || "",
              bold: cell.style?.bold || false,
              italic: cell.style?.italic || false,
              fontFace: cell.style?.typeface || "",  // font → fontFace
              align: align,  // 変換済み
              valign: valign,  // 変換済み
              margin: [
                parseFloat(emuToInch(cell.margins?.top || 0).toFixed(3)),
                parseFloat(emuToInch(cell.margins?.right || 0).toFixed(3)),
                parseFloat(emuToInch(cell.margins?.bottom || 0).toFixed(3)),
                parseFloat(emuToInch(cell.margins?.left || 0).toFixed(3))
              ],
              colspan: cell.colspan || 1,
              rowspan: cell.rowspan || 1
              // isMergedContinuation は不要（既にフィルタリング済み）
            };

            // fillをオブジェクト形式に
            if (cell.fill?.color) {
              cellData.fill = { color: cell.fill.color };
            }

            // borderをオブジェクト形式に（情報として保持、ただしPptxGenJSでは使用不可）
            cellData.border = {
              top: { pt: parseFloat(emuToPoint(cell.borders?.top?.width || 0).toFixed(2)), color: cell.borders?.top?.color || "" },
              right: { pt: parseFloat(emuToPoint(cell.borders?.right?.width || 0).toFixed(2)), color: cell.borders?.right?.color || "" },
              bottom: { pt: parseFloat(emuToPoint(cell.borders?.bottom?.width || 0).toFixed(2)), color: cell.borders?.bottom?.color || "" },
              left: { pt: parseFloat(emuToPoint(cell.borders?.left?.width || 0).toFixed(2)), color: cell.borders?.left?.color || "" }
            };

            return cellData;
          })
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
        line: {
          pt: parseFloat(emuToPoint(line.lineWidth || 0).toFixed(2)),
          color: normalizeColorHex(line.lineColor || "000000"),
          dashType: line.lineDash || "solid",
          beginArrowType: line.arrowStart !== "none" ? line.arrowStart : undefined,
          endArrowType: line.arrowEnd !== "none" ? line.arrowEnd : undefined
        }
      };
    })
  };

  // AI用の詳細プロンプトを生成
  const prompt = `# PowerPoint Slide Reproduction Task

PptxGenJSで下記のパワポを完全再現して。**位置・サイズ・色・フォント・罫線**全て完璧に。

## 重要な実装ルール

**全てのslide.addXXXメソッド（addText、addTable、addShape、addImageなど）について：**
- データ（テキストの内容、図形のプロパティ、テーブルのデータなど）を**変数に格納せず**、メソッドの引数として**直接記述**してください
- 例：
  - ❌ 間違い: \`let tableData = [[...]]; slide.addTable(tableData, {...});\`
  - ✅ 正しい: \`slide.addTable([[...]], {...});\`
- これはコードの可読性と保守性のため、変数宣言を減らし、コードを簡潔にするためです

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

各要素（**すべてPptxGenJS API形式で出力済み**）:
- **id**: ID
- **text**: テキスト内容（全段落を結合）
- **x, y**: 位置(インチ)
- **w, h**: サイズ(インチ)
- **fontSize**: フォントサイズ(pt)
- **color**: テキスト色(RGB hex, "000000"=黒)
- **bold, italic**: スタイル
- **fontFace**: フォント名（PptxGenJS形式）
- **align**: 配置("left"/"center"/"right"）（PptxGenJS形式）
- **fill**: 背景色オブジェクト { color: "FFFFFF" } または undefined
- **line**: 枠線オブジェクト { color: "000000", pt: 1.5 } または undefined
- **paragraphs**: 段落配列（箇条書き含む、存在する場合のみ）
  - **text**: 段落テキスト（全ランを結合した文字列）
  - **level**: インデントレベル(0=なし, 1以上=箇条書きレベル)
  - **bullet**: 箇条書き情報(nullは箇条書きなし)
    - **type**: "char"=文字マーカー, "number"=番号付き
    - **char**: 箇条書き文字（例: "•", "n"）
    - **font**: 箇条書きフォント（例: "Wingdings"）
    - **numType**: 番号タイプ（type="number"の場合、例: "arabicPeriod"）
  - **runs**: テキストラン配列（段落内の各テキスト断片、途中で色や太字が変わる場合に使用）
    - **text**: ランのテキスト
    - **fontSize**: フォントサイズ(pt)
    - **color**: テキスト色(RGB hex)
    - **bold**: 太字
    - **italic**: 斜体
    - **fontFace**: フォント名（PptxGenJS形式）
    - **underline**: 下線 ("none"/"sng"/"dbl")
    - **baseline**: 上付き・下付き文字 (0=通常, 正=上付き, 負=下付き)
    - **isBreak**: 改行の場合true

### Tables (表)

各表（すべてPptxGenJS API形式で出力済み）:
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
    - **fontFace**: フォント名（PptxGenJS形式）
    - **align**: 配置("left"/"center"/"right"）（PptxGenJS形式）
    - **valign**: 垂直配置("top"/"middle"/"bottom"）（PptxGenJS形式）
    - **fill**: セル背景色オブジェクト { color: "FFFFFF" } または undefined
    - **border**: 罫線オブジェクト { top: { pt: 1.0, color: "000000" }, right: {...}, bottom: {...}, left: {...} }（PptxGenJS形式、typeやdashTypeは含まない）
    - **margin**: マージン配列[top, right, bottom, left](インチ)
    - **colspan**: 列結合数(1=通常セル, 2以上=複数列にまたがる)
    - **rowspan**: 行結合数(1=通常セル, 2以上=複数行にまたがる)
    - **isMergedContinuation**: 結合継続セルフラグ(true=結合の一部でプレースホルダー, false=通常セルまたは結合主セル)

### Lines (線・コネクタ)

各線（すべてPptxGenJS API形式で出力済み）:
- **id**: ID
- **x, y**: 開始位置(インチ)
- **w, h**: 幅・高さ(インチ) ※wが線の長さ(水平), hが高さ(垂直)
- **line**: 線プロパティオブジェクト（PptxGenJS形式）
  - **pt**: 線の太さ(pt)
  - **color**: 線の色(RGB hex)
  - **dashType**: 線のスタイル("solid"=実線, "dash"=破線, "dot"=点線, "dashDot"=一点鎖線, "lgDash"=長い破線, "sysDot"=システム点線など)
  - **beginArrowType**: 開始側の矢印("arrow"=矢印, "triangle"=三角, "diamond"=菱形, "oval"=丸など、なしの場合はundefined)
  - **endArrowType**: 終了側の矢印(同上、なしの場合はundefined)

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
  // 形状タイプに応じて適切なメソッドを使用
  const shapeOptions = {
    x: el.x,
    y: el.y,
    w: el.w,
    h: el.h,
    fill: el.fill,
    line: el.line
  };

  // テキストがある場合
  if (el.text || (el.paragraphs && el.paragraphs.length > 0)) {
    // 箇条書きがある場合は段落ごとに処理
    if (el.paragraphs && el.paragraphs.length > 0) {
      const textContent = el.paragraphs.flatMap(p => {
        // runs配列がある場合は複数ラン対応
        if (p.runs && p.runs.length > 0) {
          return p.runs.map((run, runIndex) => {
            const runOptions = {
              fontSize: run.fontSize,
              color: run.color,
              bold: run.bold,
              italic: run.italic,
              fontFace: run.fontFace,
              underline: run.underline !== 'none' ? { style: run.underline } : undefined,
              subscript: run.baseline < 0,
              superscript: run.baseline > 0,
              breakLine: run.isBreak
            };

            // 最初のランのみ箇条書き設定
            if (runIndex === 0) {
              runOptions.bullet = false;
              runOptions.indentLevel = p.level;

              if (p.bullet) {
                if (p.bullet.type === 'char') {
                  runOptions.bullet = {
                    type: p.bullet.char,
                    characterCode: p.bullet.char.charCodeAt(0).toString(16)
                  };
                  if (p.bullet.font) {
                    runOptions.bullet.fontFace = p.bullet.font;
                  }
                } else if (p.bullet.type === 'number') {
                  runOptions.bullet = { type: p.bullet.numType || 'number' };
                }
              }
            }

            return { text: run.text, options: runOptions };
          });
        } else {
          // runs配列がない場合は従来の方法
          const options = {
            bullet: false,
            indentLevel: p.level
          };

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
        }
      });

      // 形状タイプに応じて適切なshapeオプションを追加
      if (el.shapeType && el.shapeType !== 'rect') {
        shapeOptions.shape = pptx.ShapeType[el.shapeType] || el.shapeType;
      }

      slide.addText(textContent, shapeOptions);
    } else {
      // 箇条書きなしの場合
      const textOptions = {
        ...shapeOptions,
        fontSize: el.fontSize,
        color: el.color,
        bold: el.bold,
        italic: el.italic,
        fontFace: el.fontFace,
        align: el.align,
        breakLine: true
      };

      // 形状タイプに応じて適切なshapeオプションを追加
      if (el.shapeType && el.shapeType !== 'rect') {
        textOptions.shape = pptx.ShapeType[el.shapeType] || el.shapeType;
      }

      slide.addText(el.text, textOptions);
    }
  } else {
    // テキストがない場合
    // homePlate, triangle などは addText("", { shape: ... }) を使用
    if (el.shapeType === 'homePlate' || el.shapeType === 'triangle') {
      const textOptions = {
        ...shapeOptions,
        fontSize: el.fontSize || 11,
        color: el.color || "000000",
        bold: el.bold || false,
        italic: el.italic || false,
        fontFace: el.fontFace || "",
        align: el.align || "center",
        shape: pptx.ShapeType[el.shapeType]
      };
      slide.addText("", textOptions);
    } else if (el.shapeType === 'ellipse') {
      // 円形は addShape を使用
      slide.addShape(pptx.ShapeType.ellipse, shapeOptions);
    } else if (el.shapeType && el.shapeType !== 'rect') {
      // その他の特殊図形も addShape を使用
      slide.addShape(pptx.ShapeType[el.shapeType] || el.shapeType, shapeOptions);
    } else {
      // デフォルトは四角形
      slide.addShape(pptx.ShapeType.rect, shapeOptions);
    }
  }
});
\`\`\`

### 表

\`\`\`javascript
tables.forEach(table => {
  const tableData = table.rows.map(row =>
    row.cells.map(cell => {
      const cellOptions = {
        fontSize: cell.fontSize,
        color: cell.color,
        bold: cell.bold,
        italic: cell.italic,
        fontFace: cell.fontFace,
        align: cell.align,
        valign: cell.valign,
        fill: cell.fill,
        border: [
          { pt: cell.border.top.pt, color: cell.border.top.color },
          { pt: cell.border.right.pt, color: cell.border.right.color },
          { pt: cell.border.bottom.pt, color: cell.border.bottom.color },
          { pt: cell.border.left.pt, color: cell.border.left.color }
        ]
        // margin: cell.margin  ← 削除: PptxGenJSのテーブルセルでmarginは使用しない
      };

      // セル結合の設定（結合セルの場合のみ）
      if (cell.colspan > 1) {
        cellOptions.colspan = cell.colspan;
      }
      if (cell.rowspan > 1) {
        cellOptions.rowspan = cell.rowspan;
      }

      return {
        text: cell.text,
        options: cellOptions
      };
    })
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
lines.forEach(lineItem => {
  slide.addShape("line", {
    x: lineItem.x,
    y: lineItem.y,
    w: lineItem.w,
    h: lineItem.h,
    line: lineItem.line
  });
});
\`\`\`

---

## Important Notes

1. **JSON形式はPptxGenJS API形式で出力済み**: プロパティ名の変換は不要、そのまま使用可能
2. **fontFace**: フォント名プロパティ（font → fontFace に変換済み）
3. **align/valign**: 配置値は変換済み（"l"/"ctr"/"r" → "left"/"center"/"right", "t"/"m"/"b" → "top"/"middle"/"bottom"）
4. **fill**: オブジェクト形式 { color: "FFFFFF" } または undefined（空文字列ではない）
5. **line**: オブジェクト形式 { color: "000000", pt: 1.5 } または undefined
6. **border**: オブジェクト形式 { top: {pt, color}, right: {pt, color}, bottom: {pt, color}, left: {pt, color} }
7. **テーブルborderについて**: PptxGenJSのテーブルでは、border配列に変換が必要（実装例参照）
8. **テーブル内のセルのmarginについて**: **絶対にmarginは使わないでください**。PptxGenJSのテーブルセルでmarginを指定すると正しく動作しません。
9. **Colors**: 6-digit RGB hex ("FF0000"=red)
10. **Units**: All positions/sizes in inches, fonts in points
11. **Header rows**: hasHeader=true means row 1 is header (already has proper background color)
12. **Line styles**: "solid", "dash", "dot", "dashDot", "lgDash", "lgDashDot", "sysDash", "sysDot"
13. **Arrow types**: "none", "arrow", "triangle", "diamond", "oval", "stealth"
14. **Line breaks**: Text contains "\\n" for line breaks. Use breakLine: true in PptxGenJS addText options
15. **Bullets**: Use paragraphs array for bullet points. level indicates indent (0=none, 1+=levels). bullet.char with bullet.font (e.g., Wingdings) for custom markers
16. **Text Runs**: When a paragraph has runs array, use it for precise styling control. Each run has its own color, bold, italic, etc. This is essential for text with mixed styles (e.g., "Complete PoC with **Ichiba & Travel** first" where "Ichiba & Travel" is red and bold)
17. **Merged Cells**:
    - **All cells are included** in the cells array (including continuation placeholders to maintain grid alignment)
    - **Primary cells** have colspan>1 or rowspan>1 and contain the actual text
    - **Continuation cells** have isMergedContinuation=true and empty text (these are placeholders)
    - In PptxGenJS, only specify colspan/rowspan on primary cell (skip cells with isMergedContinuation=true)
    - colspan: horizontal merge (columns), rowspan: vertical merge (rows)
    - Cell indices in the JSON directly correspond to grid positions in the table (maintaining alignment)

---

## Slide Data (JSON)

\`\`\`json
${JSON.stringify(data, null, 2)}
\`\`\`

---

## Data Extraction Summary

- **Elements**: ${data.elements.length} items
- **Tables**: ${data.tables.length} tables
${data.tables.map((t, i) => `  - Table ${i + 1}: ${t.rows.length} rows × ${t.colW.length} columns (${t.rows.reduce((sum, row) => sum + row.cells.length, 0)} total cells)`).join('\n')}
- **Lines**: ${data.lines.length} lines

✅ **All data extracted completely - every row, cell, and element is included above.**

---

**完璧に再現してください。位置・サイズ・色・フォントが1pxでもずれないように。**`;

  return prompt;
}

// メイン解析関数
export async function analyzePPTX(file) {
  try {
    // JSZipをロード（ブラウザ環境ではグローバル、Node.js環境ではrequire）
    let JSZipLib;
    if (typeof JSZip !== 'undefined') {
      JSZipLib = JSZip;
    } else if (typeof require !== 'undefined') {
      JSZipLib = require('jszip');
    } else {
      throw new Error('JSZip is not loaded. Please check manifest.json content_scripts configuration.');
    }

    const buf = await file.arrayBuffer();
    const zip = await JSZipLib.loadAsync(buf);

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

      // グループ化された図形を抽出
      const groupElements = await extractGroups(doc, themeColors, masterStyles, zip, slidePath);

      // グループ化されていない図形を抽出
      const elements = await extractElements(doc, themeColors, masterStyles, zip, slidePath);
      const tables = extractTables(doc, themeColors);
      const lines = extractLines(doc, themeColors);

      // グループ化された要素と非グループ化要素を結合
      const allElements = [...elements, ...groupElements.elements];
      const allTables = [...tables, ...groupElements.tables];
      const allLines = [...lines, ...groupElements.lines];

      const promptWithJson = generatePromptWithJSON(allElements, allTables, allLines, template, slidePath);

      results.push({
        slideNumber: parseInt(slidePath.match(/slide(\d+)\.xml/i)[1], 10),
        slidePath: slidePath,
        promptWithJson: promptWithJson,
        elementCount: allElements.length,
        tableCount: allTables.length,
        lineCount: allLines.length
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
