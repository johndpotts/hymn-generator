const docx = require("docx");
const fs = require("fs");
const moment = require("moment");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType
} = require("docx");
const _ = require("lodash");

const day = process.argv[2];

const monthNames = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December"
];

var d = new Date(day);

const daysInMonth = function(month, year) {
  return new Date(year, month, 0).getDate();
};

var getTot = daysInMonth(d.getMonth(), d.getFullYear()); //Get total days in a month
var sundays = new Array(); //Declaring array for inserting Sundays

const hymns = `002. Holy, Holy, Holy
004. To God Be the Glory
007. Joyful Joyful We Adore Thee
008. A Mighty Fortress Is Our God
010. How Great Thou Art
014. Praise to the Lord, the Almighty
015. Come, Thou Fount of Every Blessing
016. O Worship the King
043. This is My Father's World
054. Great is Thy Faithfulness
061. Savior, Like a Shepherd Lead Us
132. There Is Power in the Blood
135. Nothing but the Blood
136. Are You Washed in the Blood
138. At Calvary
139. At the Cross
140. Down at the Cross
141. The Old Rugged Cross
144. When I Survey the Wondrous Cross
151. The Way of the Cross Leads Home
161. Crown Him with Many Crowns
176. Fairest Lord Jesus
182. What a Friend We Have in Jesus
187. In the Garden
189. The Lily of the Valley
202. All Hail the Power of Jesus' Name
203. His Name is Wonderful
210. My Jesus, I Love Thee
216. O for a Thousand Tongues to Sing
217. Oh, How I Love Jesus
227. Praise Him! Praise Him!
247. Come, Thou Almighty King
248. God, Our Father, We Adore Thee
329. Grace Greater than Our Sin
333. Leaning on the Everlasting Arms
334. Blessed Assurance, Jesus Is Mine
335. Standing on the Promises
342. Rock of Ages, Cleft for Me
352. Faith of Our Fathers
406. The Solid Rock
410. It Is Well with My Soul
411. Tis So Sweet to Trust in Jesus
413. Faith Is the Victory
424. Heavenly Sunlight
425. He Keeps Me Singing
426. Victory In Jesus
430. Sunshine in My Soul
438. Heaven Came Down
467. There Shall Be Showers of Blessing
484. Higher Ground
485. Stand Up, Stand Up for Jesus
493. Onward, Christian Soldiers
502. Open My Eyes That I May See
514. When We All Get To Heaven
516. When the Roll Is Called Up Yonder
518. Shall We Gather at the River
524. We're Marching to Zion
537. I Will Sing the Wondrous Story
544. Redeemed, How I Love to Proclaim It
546. Love Lifted Me
547. I Stand Amazed in the Presence
581. We Have Heard the Joyful Sound
595. Send the Light
629. God of Our Fathers
633. Mine Eyes Have Seen the Glory
644. Count Your Blessings`;

const slowHymns = `275. I Surrender All
277. Take My Life, and Let It Be Consecrated
280. Jesus, Keep Me Near the Cross
294. Have Thine Own Way, Lord
307. Just As I Am
312. Softly and Tenderly
316. Jesus is Tenderly Calling
317. Only Trust Him
445. Sweet Hour of Prayer
446. Take Time to Be Holy
447. Trust and Obey
448. Just a Closer Walk with Thee
450. I Need Thee Every Hour
330. Amazing Grace! How Sweet the Sound
134. Jesus Paid It All`;

const christmasHymns = `076. O Come, O Come, Emmanuel
077. Come, Thou Long-Expected Jesus
085. The First Nowell
086. O Little Town of Bethlehem
087. Joy to the World! The Lord Is Come
088. Hark! The Herald Angels Sing
089. O Come, All Ye Faithful
091. Silent Night, Holy Night
093. It Came Upon the Midnight Clear
094. Angels, from the Realms of Glory
095. Go Tell it On The Mountain
096. Good Christian Men, Rejoice
098. I Heard the Bells on Christmas Day
100. Angels We Have Heard on High
103. Away in A Manger
108. How Great Our Joy
113. We Three Kings of Orient Are
118. What Child is This`;

const verses = [
  " Lord, you are my God; I will exalt you and praise your name, for in perfect faithfulness you have done wonderful things, things planned long ago. -Isaiah 25:1",
  "Let everything that has breath praise the Lord. Praise the Lord. -Psalm 150:6",
  "About midnight Paul and Silas were praying and singing hymns to God, and the other prisoners were listening to them. -Acts 16:25",
  "God is spirit, and his worshipers must worship in the Spirit and in truth. -John 4:24",
  "Praise the Lord, my soul; all my inmost being, praise his holy name. -Psalm 103:1",
  "Give thanks to the Lord, for he is good; his love endures forever. -1 Chronicles 16:34",
  "Though the fig tree does not bud and there are no grapes on the vines, though the olive crop fails and the fields produce no food, though there are no sheep in the pen and no cattle in the stalls, yet I will rejoice in the Lord, I will be joyful in God my Savior. -Habakkuk 3:17-18",
  "You, God, are my God, earnestly I seek you; I thirst for you, my whole being longs for you, in a dry and parched land where there is no water. -Psalm 63:1",
  "My mouth is filled with your praise, declaring your splendor all day long. -Psalm 71:8",
  "Yours, Lord, is the greatness and the power and the glory and the majesty and the splendor, for everything in heaven and earth is yours. Yours, Lord, is the kingdom; you are exalted as head over all. -1 Chronicles 29:11",
  "Praise be to the God and Father of our Lord Jesus Christ, the Father of compassion and the God of all comfort, who comforts us in all our troubles, so that we can comfort those in any trouble with the comfort we ourselves receive from God. -2 Corinthians 1:3-4",
  "How great you are, Sovereign Lord! There is no one like you, and there is no God but you, as we have heard with our own ears. -2 Samuel 7:22",
  "For from him and through him and for him are all things. To him be the glory forever! Amen. -Romans 11:36",
  "Why, my soul, are you downcast? Why so disturbed within me? Put your hope in God, for I will yet praise him, my Savior and my God. -Psalm 42:11",
  "Then you will call on me and come and pray to me, and I will listen to you. -Jeremiah 29:12",
  "It is written: ‘As surely as I live,’ says the Lord, ‘every knee will bow before me; every tongue will acknowledge God.’ -Romans 14:11",
  "Come, let us bow down in worship, let us kneel before the Lord our Maker. -Psalm 95:6 ",
  "Give praise to the Lord, proclaim his name; make known among the nations what he has done. -Psalm 105:1",
  "I spread out my hands to you; I thirst for you like a parched land. -Psalm 143:6",
  "There is no one holy like the Lord; there is no one besides you; there is no Rock like our God. -1 Samuel 2:2 ",
  "I say to the Lord, “You are my Lord; apart from you I have no good thing.” -Psalm 16:2"
];

const arr1 = hymns.split(`\n`);
let indexedHymns = arr1.map(hymn => {
  hymnArr = hymn.split(".");
  return {
    title: hymnArr[1].substring(1, hymnArr[1].length),
    number: hymnArr[0]
  };
});

const arr2 = slowHymns.split(`\n`);
let indexedSlowHymns = arr2.map(hymn => {
  hymnArr = hymn.split(".");
  return {
    title: hymnArr[1].substring(1, hymnArr[1].length),
    number: hymnArr[0]
  };
});

const arr3 = christmasHymns.split(`\n`);
let indexedChristmasHymns = arr3.map(hymn => {
  hymnArr = hymn.split(".");
  return {
    title: hymnArr[1].substring(1, hymnArr[1].length),
    number: hymnArr[0]
  };
});

const randomHymn = function() {
  let hymn = _.sample(indexedHymns);
  indexedHymns = indexedHymns.filter(hymn1 => hymn1.number !== hymn.number);
  return hymn;
};

const slowHymn = function() {
  let hymn = _.sample(indexedSlowHymns);
  indexedSlowHymns = indexedSlowHymns.filter(
    hymn1 => hymn1.number !== hymn.number
  );
  return hymn;
};

const christmasHymn = function() {
  let hymn = _.sample(indexedChristmasHymns);
  indexedChristmasHymns = indexedChristmasHymns.filter(
    hymn1 => hymn1.number !== hymn.number
  );
  return hymn;
};

for (var i = 1; i <= getTot; i++) {
  //looping through days in month
  var newDate = new Date(d.getFullYear(), d.getMonth(), i);
  if (newDate.getDay() == 0) {
    //if Sunday
    sundays.push(newDate);
  }
}

const doc = new Document();
const arr = new Array();
sundays.forEach(sunday => {
  arr.push(
    new Paragraph({
      children: [
        new TextRun({
          text: `${moment(sunday).format("MM/DD/YYYY")}`,
          underline: {}
        })
      ]
    })
  );
  arr.push(
    new Paragraph({
      children: [
        new TextRun({
          text: ` `
        })
      ]
    })
  );
  let hymn = d.getMonth() === 11 ? christmasHymn() : randomHymn();
  arr.push(
    new Paragraph({
      children: [
        new TextRun({
          text: `${hymn.title} - ${hymn.number}`
        })
      ]
    })
  );
  arr.push(
    new Paragraph({
      children: [
        new TextRun({
          text: ` `
        })
      ]
    })
  );
  hymn = d.getMonth() === 11 ? christmasHymn() : randomHymn();
  arr.push(
    new Paragraph({
      children: [
        new TextRun({
          text: `${hymn.title} - ${hymn.number}`
        })
      ]
    })
  );
  arr.push(
    new Paragraph({
      children: [
        new TextRun({
          text: ` `
        })
      ]
    })
  );
  hymn = d.getMonth() === 11 ? christmasHymn() : slowHymn();

  arr.push(
    new Paragraph({
      children: [
        new TextRun({
          text: `${hymn.title} - ${hymn.number}`
        })
      ]
    })
  );
  arr.push(
    new Paragraph({
      children: [
        new TextRun({
          text: ` `
        })
      ]
    })
  );
  hymn = slowHymn();

  arr.push(
    new Paragraph({
      children: [
        new TextRun({
          text: `Response: ${hymn.title} - ${hymn.number}`
        })
      ]
    })
  );
  arr.push(
    new Paragraph({
      children: [
        new TextRun({
          text: ` `
        })
      ]
    })
  );
  arr.push(
    new Paragraph({
      children: [
        new TextRun({
          text: ` `
        })
      ]
    })
  );
  arr.push(
    new Paragraph({
      children: [
        new TextRun({
          text: ` `
        })
      ]
    })
  );
});

doc.addSection({
  children: [
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: `Hymns for ${moment(d).format("MMMM YYYY")} `
        })
      ]
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: ` `
        })
      ]
    }),
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      alignment: AlignmentType.CENTER,

      children: [
        new TextRun({
          text: `${verses[d.getMonth()]} `,
          italics: true
        })
      ]
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: ` `
        })
      ]
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: ` `
        })
      ]
    }),
    ...arr
  ]
});

const packer = new Packer();

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(
    `${monthNames[d.getMonth()]} ${d.getFullYear()} hymns.docx`,
    buffer
  );
});
