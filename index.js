var pptxgen = require('pptxgenjs');
    let pptx = new pptxgen();
    // Set Layout
    pptx.defineLayout({ name:'PenolakanVaskin', width:13.333, height:7.5 });
    pptx.layout = "PenolakanVaskin"

    // 2. Add a Slide
    let slide = pptx.addSlide();
    let horizontal_position = 1.35 , vertical_position = 0.24, width = 7.82, height = 0.78

    let textboxOpts = { 
      x: horizontal_position,
      y: vertical_position, 
      w: width, 
      h: height,
      color: "7F7F7F",
      fontSize: 36 };

    let textboxText = [
      {
        text: "Exposure",
        options: {}
      },
      {
        text: " \"Penolakan Vaksin\"",
        options: {
          bold:true
        }
      }
    ]
    slide.addText(textboxText, textboxOpts);


    let textboxOpts2 = { 
      x: 4.17,
      y: 0.91, 
      w: 3.84, 
      h: 0.4,
      color: "D84E2E",
      fontSize: 18 };

    let textboxText2 = [
      {
        text: "1 Januari 2021 – 31 Maret 2021",
        options: {}
      }
    ]

    slide.addText(textboxText2, textboxOpts2);

    let shapesOption = { 
      h:0.36, 
      x:8.19, 
      y:0.28,
      w: 4.88,
      fill: {
        color: "EBB55A", 
      }
    } 

    slide.addShape(pptx.ShapeType.rect, shapesOption);

    let shapesOption2 = {
      
      w: 3.63, 
      h:7.28, 
      x:9.43, 
      y:0.21,
      fill: { 
        color: "EBB55A",
        transparency: 71
      }
  }
    slide.addShape(pptx.ShapeType.rect, shapesOption2);
    // 4. Save the Presentation

    let textboxText3 = [
      {
        text: "Pembahasan Vaksin mulai ramai dibicarakan netizen pada bulan Januari (sejak wacana vaksin dimulai),",
        options: {
          bold:true,
          bullet: true
        }
      },
      {
        text: " di mana pemerintah mengharuskan seluruh masyarakat untuk mengikuti vaksinasi, namun diikuti dengan pernyataan dari",
        options: {
        }
      },
      {
        text: " Anggota DPR Ribka Tjiptaning dalam forum resmi legislatif yang menyatakan menolak menerima vaksin corona buatan perusahaan farmasi asal China, Sinovac.",
        options: {
          bold:true,
          paraSpaceBefore: 12,
          paraSpaceAfter: 24,
        }
      },
      {
        text: "Vaksinasi Covid-19 tahap pertama mulai bergulir di berbagai daerah walau sejumlah kalangan masih enggan mengikuti salah satu upaya mengatasi pandemi covid-19.",
        options: {
          bullet: true,
          paraSpaceBefore: 7,
        }
      },
      {
        text: "Trend pembicaraan penolakan vaksin terus bergerak menurun. Sementara tagar #TolakDivaksinSinovac sempat mencuat di Twitter karena dicuitkan belasan ribu kali.",
        options: {
          bold:true,
        }
      },
    ]

    let textboxOpts3 = { 
      x: 0.38,
      y: 5.14, 
      w: 8.71, 
      h: 2.06,
      wrap:	true,
      fontSize: 14 
    };
    
    let dataChartPie = [
        {
          name: "Sales",
          labels: ["Jumlah reaksi pembicaraan vaksin secara keseluruhan","jumlah reaksi penolakan vaksin"],
          values: [83.797,34.854],
        }
      ];

    let chartOption = { 
      x: 9.41,
      y: 0.91, 
      w: 3.66, 
      h: 2.97,
      dataLabelColor: "FFFFFF",
      chartColors:["D84E2E","7E979B"],
      legendPos: 'b',
      showLegend: true,
      legendFontSize: 12,
      dataLabelFontSize: 16,
      dataLabelFontBold: true
     }

     let dataChartPie2 = [
      {
        name: "Sales",
        labels: ["Jumlah akun yang membicarakan vaksin secara keseluruhan","Jumlah akun penolak Vaksin"],
        values: [643855,222250],
      }
    ];

    let chartOption2 = { 
      x: 9.31,
      y: 4.63, 
      w: 3.66, 
      h: 2.5,
      dataLabelColor: "FFFFFF",
      chartColors:["D84E2E","7E979B"],
      legendPos: 'b',
      showLegend: true,
      legendFontSize: 9,
      dataLabelFontSize: 12,
      dataLabelFontBold: true,
      layout:{
      x: 0.3,
      y: 0.2, 
      w: .5, 
      h: .5,
      }
    }

    const comboTypes = [
      {
        type: pptx.charts.BAR,
        data: [{
          name: "Post",
          labels: ["Jan-21","Feb-21","Mar-21"],
          values: [29173,3534,3314],
        }
      ],
        options:{
          showLabel: true,
          chartColors: ["D84E2E"],
          dataLabelColor: "000000",
          dataLabelFontFace: "Arial",
          dataLabelFontSize: 9,
          dataLabelPosition: "outEnd",
          showValue: true,
        }
      },
      {
        type: pptx.charts.LINE,
        data: [{
              name: "Akun",
              labels: ["Jan-21","Feb-21","Mar-21"],
              values: [17004,2705,2541]
            },
          ],
        options: {
          lineSmooth : true,
          lineDataSymbol: "none",
          chartColors:["7E979B"]
        }
      },
    ];

    let comboOption = { 
      x: 0.32,
      y: 1.3, 
      w: 8.51, 
      h: 3.61,
      legendPos: 'l',
      showLegend: true,
    }

    slide.addText(textboxText3, textboxOpts3);
    slide.addChart(pptx.ChartType.pie, dataChartPie,chartOption);
    slide.addChart(pptx.ChartType.pie, dataChartPie2,chartOption2);
    slide.addChart(comboTypes,comboOption);

    pptx.writeFile({fileName:"Contoh.pptx"});