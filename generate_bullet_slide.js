const PptxGenJS = require("pptxgenjs");

// 提示されたJSONデータ
const slideData = {
  "slide": "slide1",
  "template": {
    "background": "",
    "defaultTextColor": "",
    "slideNumber": {
      "x": 12.671,
      "y": 7.044,
      "w": 0.569,
      "h": 0.252,
      "fontSize": 9,
      "font": "",
      "color": "000000",
      "bold": true,
      "align": "l"
    },
    "fixedImages": []
  },
  "elements": [
    {
      "id": 1,
      "text": "Items to align on about Fusion PoC",
      "x": 0.5,
      "y": 0.4,
      "w": 12.33,
      "h": 0.8,
      "fontSize": 22,
      "color": "000000",
      "bold": true,
      "italic": false,
      "fontFace": "Rakuten Sans JP-2",
      "align": "left",
      "shapeType": "rect",
      "paragraphs": [
        {
          "text": "Items to align on about Fusion PoC",
          "level": 0,
          "bullet": null,
          "runs": [
            {
              "text": "Items to align on about Fusion PoC",
              "fontSize": 22,
              "color": "000000",
              "bold": true,
              "italic": false,
              "fontFace": "Rakuten Sans JP-2",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            }
          ]
        }
      ]
    },
    {
      "id": 2,
      "text": "",
      "x": 0.273,
      "y": 1.217,
      "w": 6.2,
      "h": 2.938,
      "fontSize": 18,
      "color": "000000",
      "bold": false,
      "italic": false,
      "fontFace": "",
      "align": "left",
      "shapeType": "rect",
      "fill": {
        "color": "F7FAFC"
      },
      "line": {
        "color": "E2E8F0",
        "pt": 0.5
      },
      "paragraphs": [
        {
          "text": "",
          "level": 0,
          "bullet": null,
          "runs": []
        }
      ]
    },
    {
      "id": 3,
      "text": "1. Overall Direction",
      "x": 0.573,
      "y": 1.267,
      "w": 5.5,
      "h": 0.35,
      "fontSize": 16,
      "color": "1A202C",
      "bold": true,
      "italic": false,
      "fontFace": "Rakuten Sans JP-2",
      "align": "left",
      "shapeType": "rect",
      "paragraphs": [
        {
          "text": "1. Overall Direction",
          "level": 0,
          "bullet": null,
          "runs": [
            {
              "text": "1. Overall Direction",
              "fontSize": 16,
              "color": "1A202C",
              "bold": true,
              "italic": false,
              "fontFace": "Rakuten Sans JP-2",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            }
          ]
        }
      ]
    },
    {
      "id": 4,
      "text": "Complete PoC with Ichiba & Travel first\n\nBegin PoC with Mart, Beauty, Keiba/Kdreams, J-League thereafter\n\nIf necessary, Builder can extend the PoC accounts past Nov.\n\n",
      "x": 0.273,
      "y": 1.717,
      "w": 6.2,
      "h": 2.427,
      "fontSize": 14,
      "color": "4A5568",
      "bold": false,
      "italic": false,
      "fontFace": "Rakuten Sans JP",
      "align": "left",
      "shapeType": "rect",
      "paragraphs": [
        {
          "text": "Complete PoC with Ichiba & Travel first",
          "level": 1,
          "bullet": {
            "type": "char",
            "char": "•",
            "font": "Arial"
          },
          "runs": [
            {
              "text": "Complete PoC with ",
              "fontSize": 14,
              "color": "4A5568",
              "bold": false,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            },
            {
              "text": "Ichiba & Travel",
              "fontSize": 14,
              "color": "BF0000",
              "bold": true,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            },
            {
              "text": " first",
              "fontSize": 14,
              "color": "4A5568",
              "bold": false,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            }
          ]
        },
        {
          "text": "",
          "level": 1,
          "bullet": {
            "type": "char",
            "char": "•",
            "font": "Arial"
          },
          "runs": []
        },
        {
          "text": "Begin PoC with Mart, Beauty, Keiba/Kdreams, J-League thereafter",
          "level": 1,
          "bullet": {
            "type": "char",
            "char": "•",
            "font": "Arial"
          },
          "runs": [
            {
              "text": "Begin PoC with Mart, Beauty, Keiba/",
              "fontSize": 14,
              "color": "4A5568",
              "bold": false,
              "italic": false,
              "fontFace": "Rakuten Sans JP-2",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            },
            {
              "text": "Kdreams",
              "fontSize": 14,
              "color": "4A5568",
              "bold": false,
              "italic": false,
              "fontFace": "Rakuten Sans JP-2",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            },
            {
              "text": ", J-League thereafter",
              "fontSize": 14,
              "color": "4A5568",
              "bold": false,
              "italic": false,
              "fontFace": "Rakuten Sans JP-2",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            }
          ]
        },
        {
          "text": "",
          "level": 1,
          "bullet": {
            "type": "char",
            "char": "•",
            "font": "Arial"
          },
          "runs": []
        },
        {
          "text": "If necessary, Builder can extend the PoC accounts past Nov.",
          "level": 1,
          "bullet": {
            "type": "char",
            "char": "•",
            "font": "Arial"
          },
          "runs": [
            {
              "text": "If necessary, Builder can extend the PoC accounts ",
              "fontSize": 14,
              "color": "4A5568",
              "bold": false,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            },
            {
              "text": "past Nov.",
              "fontSize": 14,
              "color": "BF0000",
              "bold": true,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            }
          ]
        },
        {
          "text": "",
          "level": 0,
          "bullet": null,
          "runs": []
        },
        {
          "text": "",
          "level": 1,
          "bullet": {
            "type": "char",
            "char": "•",
            "font": "Arial"
          },
          "runs": []
        }
      ]
    },
    {
      "id": 5,
      "text": "",
      "x": 6.903,
      "y": 1.217,
      "w": 6.2,
      "h": 3.394,
      "fontSize": 18,
      "color": "000000",
      "bold": false,
      "italic": false,
      "fontFace": "",
      "align": "left",
      "shapeType": "rect",
      "fill": {
        "color": "F7FAFC"
      },
      "line": {
        "color": "E2E8F0",
        "pt": 0.5
      },
      "paragraphs": [
        {
          "text": "",
          "level": 0,
          "bullet": null,
          "runs": []
        }
      ]
    },
    {
      "id": 6,
      "text": "2. PoC Schedule",
      "x": 7.203,
      "y": 1.267,
      "w": 5.5,
      "h": 0.35,
      "fontSize": 16,
      "color": "1A202C",
      "bold": true,
      "italic": false,
      "fontFace": "Rakuten Sans JP-2",
      "align": "left",
      "shapeType": "rect",
      "paragraphs": [
        {
          "text": "2. PoC Schedule",
          "level": 0,
          "bullet": null,
          "runs": [
            {
              "text": "2. PoC Schedule",
              "fontSize": 16,
              "color": "1A202C",
              "bold": true,
              "italic": false,
              "fontFace": "Rakuten Sans JP-2",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            }
          ]
        }
      ]
    },
    {
      "id": 7,
      "text": "\n11/3-7: Ichiba, Travel\n\n11/10-21: Kick-off & repository connection\n\n11/24-12/5: Ichiba & Travel PoC with Builder support\nSelf-serve plan for Rakuten created in parallel\n\n12/1-12: PoC with Mart, Beauty, Keiba/Kdreams, J-league\n\n12/17-21: Collect results & report\n",
      "x": 6.889,
      "y": 1.647,
      "w": 6.214,
      "h": 2.964,
      "fontSize": 14,
      "color": "BF0000",
      "bold": true,
      "italic": false,
      "fontFace": "Rakuten Sans JP",
      "align": "left",
      "shapeType": "rect",
      "paragraphs": [
        {
          "text": "",
          "level": 1,
          "bullet": {
            "type": "char",
            "char": "•",
            "font": "Arial"
          },
          "runs": []
        },
        {
          "text": "11/3-7: Ichiba, Travel",
          "level": 1,
          "bullet": {
            "type": "char",
            "char": "•",
            "font": "Arial"
          },
          "runs": [
            {
              "text": "11/3-7:",
              "fontSize": 14,
              "color": "BF0000",
              "bold": true,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            },
            {
              "text": " Ichiba, Travel",
              "fontSize": 14,
              "color": "4A5568",
              "bold": false,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            }
          ]
        },
        {
          "text": "",
          "level": 1,
          "bullet": {
            "type": "char",
            "char": "•",
            "font": "Arial"
          },
          "runs": []
        },
        {
          "text": "11/10-21: Kick-off & repository connection",
          "level": 1,
          "bullet": {
            "type": "char",
            "char": "•",
            "font": "Arial"
          },
          "runs": [
            {
              "text": "11/10-21:",
              "fontSize": 14,
              "color": "BF0000",
              "bold": true,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            },
            {
              "text": " Kick-off & repository connection",
              "fontSize": 14,
              "color": "4A5568",
              "bold": false,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            }
          ]
        },
        {
          "text": "",
          "level": 1,
          "bullet": {
            "type": "char",
            "char": "•",
            "font": "Arial"
          },
          "runs": []
        },
        {
          "text": "11/24-12/5: Ichiba & Travel PoC with Builder support\nSelf-serve plan for Rakuten created in parallel",
          "level": 1,
          "bullet": {
            "type": "char",
            "char": "•",
            "font": "Arial"
          },
          "runs": [
            {
              "text": "11/24-12/5:",
              "fontSize": 14,
              "color": "BF0000",
              "bold": true,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            },
            {
              "text": " ",
              "fontSize": 14,
              "color": "4A5568",
              "bold": false,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            },
            {
              "text": "Ichiba",
              "fontSize": 14,
              "color": "4A5568",
              "bold": false,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            },
            {
              "text": " & Travel PoC with Builder support",
              "fontSize": 14,
              "color": "4A5568",
              "bold": false,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            },
            {
              "text": "\n",
              "fontSize": 18,
              "color": "000000",
              "bold": false,
              "italic": false,
              "fontFace": "",
              "underline": "none",
              "baseline": 0,
              "isBreak": true
            },
            {
              "text": "Self-serve plan for Rakuten created in parallel",
              "fontSize": 14,
              "color": "718096",
              "bold": false,
              "italic": false,
              "fontFace": "Rakuten Sans JP-2",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            }
          ]
        },
        {
          "text": "",
          "level": 1,
          "bullet": {
            "type": "char",
            "char": "•",
            "font": "Arial"
          },
          "runs": []
        },
        {
          "text": "12/1-12: PoC with Mart, Beauty, Keiba/Kdreams, J-league",
          "level": 1,
          "bullet": {
            "type": "char",
            "char": "•",
            "font": "Arial"
          },
          "runs": [
            {
              "text": "12/1-12:",
              "fontSize": 14,
              "color": "BF0000",
              "bold": true,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            },
            {
              "text": " PoC with Mart, Beauty, Keiba/",
              "fontSize": 14,
              "color": "4A5568",
              "bold": false,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            },
            {
              "text": "Kdreams",
              "fontSize": 14,
              "color": "4A5568",
              "bold": false,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            },
            {
              "text": ", J-league",
              "fontSize": 14,
              "color": "4A5568",
              "bold": false,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            }
          ]
        },
        {
          "text": "",
          "level": 1,
          "bullet": {
            "type": "char",
            "char": "•",
            "font": "Arial"
          },
          "runs": []
        },
        {
          "text": "12/17-21: Collect results & report",
          "level": 1,
          "bullet": {
            "type": "char",
            "char": "•",
            "font": "Arial"
          },
          "runs": [
            {
              "text": "12/17-21:",
              "fontSize": 14,
              "color": "BF0000",
              "bold": true,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            },
            {
              "text": " Collect results & report",
              "fontSize": 14,
              "color": "4A5568",
              "bold": false,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            }
          ]
        },
        {
          "text": "",
          "level": 1,
          "bullet": {
            "type": "char",
            "char": "•",
            "font": "Arial"
          },
          "runs": []
        }
      ]
    },
    {
      "id": 8,
      "text": "",
      "x": 0.273,
      "y": 4.611,
      "w": 6.2,
      "h": 2.177,
      "fontSize": 18,
      "color": "000000",
      "bold": false,
      "italic": false,
      "fontFace": "",
      "align": "left",
      "shapeType": "rect",
      "fill": {
        "color": "F7FAFC"
      },
      "line": {
        "color": "E2E8F0",
        "pt": 0.5
      },
      "paragraphs": [
        {
          "text": "",
          "level": 0,
          "bullet": null,
          "runs": []
        }
      ]
    },
    {
      "id": 9,
      "text": "3. Support",
      "x": 0.573,
      "y": 4.792,
      "w": 5.5,
      "h": 0.35,
      "fontSize": 16,
      "color": "1A202C",
      "bold": true,
      "italic": false,
      "fontFace": "Rakuten Sans JP-2",
      "align": "left",
      "shapeType": "rect",
      "paragraphs": [
        {
          "text": "3. Support",
          "level": 0,
          "bullet": null,
          "runs": [
            {
              "text": "3. Support",
              "fontSize": 16,
              "color": "1A202C",
              "bold": true,
              "italic": false,
              "fontFace": "Rakuten Sans JP-2",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            }
          ]
        }
      ]
    },
    {
      "id": 10,
      "text": "Ichiba & Travel: Hands-on support from Builder\n\nOthers:\n• Hands-on support if committed to paid account\n• Otherwise: Self-serve with OMID support\n• Builder available for specific questions\n\n",
      "x": 0.486,
      "y": 5.323,
      "w": 5.5,
      "h": 1.296,
      "fontSize": 14,
      "color": "1A202C",
      "bold": true,
      "italic": false,
      "fontFace": "Rakuten Sans JP",
      "align": "left",
      "shapeType": "rect",
      "paragraphs": [
        {
          "text": "Ichiba & Travel: Hands-on support from Builder",
          "level": 0,
          "bullet": null,
          "runs": [
            {
              "text": "Ichiba & Travel: ",
              "fontSize": 14,
              "color": "1A202C",
              "bold": true,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            },
            {
              "text": "Hands-on support",
              "fontSize": 14,
              "color": "BF0000",
              "bold": true,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            },
            {
              "text": " from Builder",
              "fontSize": 14,
              "color": "4A5568",
              "bold": false,
              "italic": false,
              "fontFace": "Rakuten Sans JP",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            }
          ]
        },
        {
          "text": "",
          "level": 0,
          "bullet": null,
          "runs": []
        },
        {
          "text": "Others:",
          "level": 0,
          "bullet": null,
          "runs": [
            {
              "text": "Others:",
              "fontSize": 14,
              "color": "1A202C",
              "bold": true,
              "italic": false,
              "fontFace": "Rakuten Sans JP-2",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            }
          ]
        },
        {
          "text": "• Hands-on support if committed to paid account",
          "level": 0,
          "bullet": null,
          "runs": [
            {
              "text": "• Hands-on support if committed to paid account",
              "fontSize": 14,
              "color": "4A5568",
              "bold": false,
              "italic": false,
              "fontFace": "Rakuten Sans JP-2",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            }
          ]
        },
        {
          "text": "• Otherwise: Self-serve with OMID support",
          "level": 0,
          "bullet": null,
          "runs": [
            {
              "text": "• Otherwise: Self-serve with OMID support",
              "fontSize": 14,
              "color": "4A5568",
              "bold": false,
              "italic": false,
              "fontFace": "Rakuten Sans JP-2",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            }
          ]
        },
        {
          "text": "• Builder available for specific questions",
          "level": 0,
          "bullet": null,
          "runs": [
            {
              "text": "• Builder available for specific questions",
              "fontSize": 14,
              "color": "4A5568",
              "bold": false,
              "italic": false,
              "fontFace": "Rakuten Sans JP-2",
              "underline": "none",
              "baseline": 0,
              "isBreak": false
            }
          ]
        },
        {
          "text": "",
          "level": 0,
          "bullet": null,
          "runs": []
        },
        {
          "text": "",
          "level": 0,
          "bullet": null,
          "runs": []
        }
      ]
    }
  ],
  "tables": [],
  "lines": []
};

// PowerPointを生成
const pptx = new PptxGenJS();
const slide = pptx.addSlide();

// テンプレート設定
if (slideData.template.background) {
  slide.background = { color: slideData.template.background };
}

// スライド番号
if (slideData.template.slideNumber) {
  slide.slideNumber = {
    x: slideData.template.slideNumber.x,
    y: slideData.template.slideNumber.y,
    fontFace: slideData.template.slideNumber.font || "Arial",
    fontSize: slideData.template.slideNumber.fontSize,
    color: slideData.template.slideNumber.color,
    bold: slideData.template.slideNumber.bold
  };
}

// 要素を追加
slideData.elements.forEach(el => {
  const shapeOptions = {
    x: el.x,
    y: el.y,
    w: el.w,
    h: el.h,
    fill: el.fill,
    line: el.line
  };

  // テキストがある場合
  if (el.paragraphs && el.paragraphs.length > 0) {
    const textContent = el.paragraphs.flatMap(p => {
      // runs配列がある場合（空配列も含む）
      if (p.runs !== undefined) {
        // 空のruns配列の場合、段落レベルで箇条書きを設定
        if (p.runs.length === 0) {
          // 空の段落は改行として扱い、箇条書きプロパティは付けない
          return { text: "", options: { breakLine: true } };
        }

        // runs配列に要素がある場合
        // ⚠️ 重要: 段落内の複数runを正しく処理するために：
        // - 最初のrunにのみbulletプロパティを設定
        // - 2つ目以降のrunはbreakLine: falseを明示的に設定
        return p.runs.map((run, runIndex) => {
          const runOptions = {
            fontSize: run.fontSize,
            color: run.color,
            bold: run.bold,
            italic: run.italic,
            fontFace: run.fontFace || "Arial"
          };

          // breakLineの処理: isBreakがtrueの場合のみ改行
          if (run.isBreak) {
            runOptions.breakLine = true;
          }

          // 最初のrunにのみ箇条書きプロパティを設定
          if (runIndex === 0) {
            if (p.bullet) {
              runOptions.bullet = true;
              runOptions.indentLevel = p.level || 0;
            } else {
              runOptions.bullet = false;
              runOptions.indentLevel = 0;
            }
          }

          return { text: run.text, options: runOptions };
        });
      } else {
        // runs配列がない場合
        const options = {
          fontSize: el.fontSize,
          color: el.color,
          bold: el.bold,
          italic: el.italic,
          fontFace: el.fontFace || "Arial",
          bullet: p.bullet ? true : false,
          indentLevel: p.level || 0
        };

        return { text: p.text, options };
      }
    });

    slide.addText(textContent, shapeOptions);
  } else {
    // テキストがない場合（背景ボックスなど）
    slide.addShape(pptx.ShapeType.rect, shapeOptions);
  }
});

// ファイルを保存
pptx.writeFile({ fileName: "箇条書きと途中で赤_再現.pptx" })
  .then(() => {
    console.log("✅ PowerPointファイルが正常に生成されました: 箇条書きと途中で赤_再現.pptx");
  })
  .catch(err => {
    console.error("❌ エラー:", err);
  });
