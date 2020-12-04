const {
  existsSync,
  mkdirSync,
  readFile,
  readFileSync,
  writeFile,
  writeFileSync,
} = require('fs')
const { resolve } = require('path')
const readline = require('readline')
require('colors')
const { google } = require('googleapis')
const { prompt } = require('inquirer')
const { sendEmail } = require('nodejs-nodemailer-outlook')
const pdfMerge = require('pdfmerge')
const PZ = require('pizzip')
const DocxTemp = require('docxtemplater')
const data = require('./data.json')

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
const TOKEN_PATH = 'token.json'

const courseQuestion = {
  type: 'list',
  name: 'course',
  message: 'Choose one of the following courses:',
  choices: ['in608', 'in721'],
}

const processQuestion = {
  type: 'list',
  name: 'process',
  message: 'Choose one of the following processes:',
  choices: ['generate pdf', 'email pdf', 'merge pdf'],
}

const { in608, in721, email, password } = data

let courseName = ''
let outputDirectory = ''
let spreadsheetId = ''
let range = ''

prompt(courseQuestion).then((answer) => {
  const { course } = answer
  outputDirectory = course
  if (!existsSync(course)) mkdirSync(course)
  switch (course) {
    case 'in608':
      courseName = 'IN608: Intermediate Application Development Concepts'
      spreadsheetId = in608.spreadsheet_id
      range = in608.practicals_range
      break
    case 'in721':
      courseName = 'IN721: Design and Development of Applications for Mobile Devices'
      spreadsheetId = in721.spreadsheet_id
      range = in721.overall_range
      break
  }
  readFile('credentials.json', (err, content) => {
    if (err) return console.log(`Error loading client secret file: ${err}`)
    authorize(JSON.parse(content), runProcess)
  })
})

const authorize = (credentials, callback) => {
  const { client_secret, client_id, redirect_uris } = credentials.installed
  const oAuth2Client = new google.auth.OAuth2(
    client_id,
    client_secret,
    redirect_uris[0]
  )

  readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client, callback)
    oAuth2Client.setCredentials(JSON.parse(token))
    callback(oAuth2Client)
  })
}

const getNewToken = (oAuth2Client, callback) => {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  })
  console.log(`Authorize this app by visiting this url: ${authUrl}`)
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  })
  rl.question('Enter the code from that page here: ', (code) => {
    rl.close()
    oAuth2Client.getToken(code, (err, token) => {
      if (err)
        return console.error(
          `Error while trying to retrieve access token ${err}`
        )
      oAuth2Client.setCredentials(token)
      writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) return console.error(err)
        console.log(`Token stored to ${TOKEN_PATH}`)
      })
      callback(oAuth2Client)
    })
  })
}

const runProcess = (auth) => {
  const sheets = google.sheets({ version: 'v4', auth })
  sheets.spreadsheets.values.get(
    {
      spreadsheetId: spreadsheetId,
      range: range,
    },
    (err, res) => {
      if (err) return console.log(`The API returned an error: ${err}`)
      const rows = res.data.values
      const studentData = []
      if (rows.length) {
        rows.map((row) => {
          // let obj = {
          //   course_name: courseName,
          //   person_code: row[0],
          //   first_name: row[1],
          //   last_name: row[2],
          //   email_address: row[3],
          //   overall_percentage: row[5],
          //   overall_grade: row[6],
          //   crit_one: row[7],
          //   crit_one_mark: row[7] * 4.5,
          //   crit_two: row[8],
          //   crit_two_mark: row[8] * 4.5,
          //   crit_three: row[9],
          //   crit_three_mark: row[9] * 1,
          //   comment_one: row[10],
          //   comment_two: row[11],
          //   comment_three: row[12]
          // }
          // let obj = {
          //   course_name: courseName,
          //   person_code: row[0],
          //   first_name: row[1],
          //   last_name: row[2],
          //   email_address: row[3],
          //   overall_percentage: row[4],
          //   overall_grade: row[5],
          //   practicals_percentage: row[6],
          //   projects_percentage: row[7],
          //   practicals_total: (row[6] * 0.2).toFixed(2),
          //   projects_total: (row[7] * 0.8).toFixed(2),
          // }
          // let obj = {
          //   course_name: courseName,
          //   person_code: row[2],
          //   first_name: row[0],
          //   last_name: row[1],
          //   email_address: row[3],
          //   overall_percentage: row[4],
          //   overall_grade: row[5],
          //   exam_percentage: row[6],
          //   software_percentage: row[7],
          //   exam_total: (row[6] * 0.3).toFixed(2),
          //   software_total: (row[7] * 0.7).toFixed(2),
          // }
          // let obj = {
          //   course_name: courseName,
          //   first_name: row[0],
          //   last_name: row[1],
          //   person_code: row[2],
          //   email_address: row[3],
          //   overall_percentage: row[5],
          //   overall_grade: row[6],
          //   crit_one: row[7],
          //   crit_one_mark: row[7] * 4,
          //   crit_two: row[8],
          //   crit_two_mark: row[8] * 5,
          //   crit_three: row[9],
          //   crit_three_mark: row[9] * 1,
          //   comment_one: row[10],
          //   comment_two: row[11],
          //   comment_three: row[12]
          // }
          // let obj = {
          //   course_name: courseName,
          //   person_code: row[0],
          //   first_name: row[1],
          //   last_name: row[2],
          //   points: row[3],
          //   percentage: row[4],
          //   grade: row[5],
          // }
          // let obj = {
          //   course_name: courseName,
          //   first_name: row[0],
          //   last_name: row[1],
          //   email_address: row[2],
          //   person_code: row[3],
          //   overall_percentage: row[4],
          //   overall_grade: row[5],
          //   practicals_percentage: row[6],
          //   projects_percentage: row[7],
          // }
          studentData.push(obj)
          // if (courseName == 'IN628 Programming 4') {
          //   obj.checkpoint_percentage = row[6]
          //   obj.software_percentage = row[8]
          //   obj.exam_percentage = row[10]
          //   studentData.push(obj)
          // } else {
          //   obj.exam_percentage = row[6]
          //   obj.software_percentage = row[8]
          //   studentData.push(obj)
          // }
        })
      } else {
        console.log('No data found.')
      }

      prompt(processQuestion).then((answer) => {
        const { process } = answer
        switch (process) {
          case 'generate pdf':
            generateDoc(studentData)
            break
          case 'email pdf':
            emailPDF(studentData)
            break
          case 'merge pdf':
            mergePDF(studentData)
            break
        }
      })
    }
  )
}

const generateDoc = (studentData) => {
  const content = readFileSync(
    resolve(__dirname, 'assessment-mobile-overall.docx'),
    'binary'
  )
  const zip = new PZ(content)
  const doc = new DocxTemp(zip)
  doc.setData(studentData)
  studentData.map((data) => {
    const firstName = data.first_name.toLowerCase()
    const lastName = data.last_name.toLowerCase()
    doc.setData(data)
    doc.render()
    const buffer = doc.getZip().generate({ type: 'nodebuffer' })
    console.log(`Generating file for ${firstName} ${lastName}.`.green)
    writeFileSync(
      resolve(__dirname, 'in608', `${firstName}-${lastName}-assessment-react-native-hacker-news-app.docx`),
      buffer
    )
    console.log(`File generated for ${firstName} ${lastName}.`.blue)
  })
}

const emailPDF = (studentData) => {
  let interval = 8000
  studentData.map((data, idx) => {
    const firstName = data.first_name.toLowerCase()
    const lastName = data.last_name.toLowerCase()
    const filename = `./${outputDirectory}/${outputDirectory}-${firstName}-${lastName}-overall-results.pdf`
    setTimeout((_) => {
      console.log(`Emailing PDF file to ${firstName} ${lastName}.`.green)
      sendEmail({
        auth: {
          user: email,
          pass: password,
        },
        from: email,
        to: data.email_address.toLowerCase(),
        // to: 'graysono@op.ac.nz',
        subject: `${courseName} practical, Django/React application & overall results`,
        html: `Kia ora ${data.first_name}, <br /> <br />
        Firstly, Tom & I would like to thank you for your efforts this semester. It has been one hell of a year but we all got through it together. I hope you have learnt something from this course. We have attached your results for the practicals, Django/React application & course. <b>Note:</b> all three results are in one PDF file. Your marks are currently provisional until they are officially released on Friday. Once released, you will be able to view your results on your Student Hub results & rewards tab. If you have any questions about your results, please do not hesitate to ask. Stay safe & enjoy your well earned break. See you next year. <br /> <br />
        Ngā mihi nui, <br /> <br />
        Grayson Orr & Tom Clark`,
        attachments: [
          {
            path: filename,
          },
        ],
        onError: (err) => console.log(err),
        onSuccess: (_) => {
          console.log(`PDF file emailed to ${firstName} ${lastName}.`.blue)
        },
      })
    }, idx * interval)
  })
}

const mergePDF = (studentData) => {
  let interval = 7500
  studentData.map((data, idx) => {
    const firstName = data.first_name.toLowerCase()
    const lastName = data.last_name.toLowerCase()
    setTimeout((_) => {
      console.log(`Merging PDF file for ${firstName} ${lastName}.`.green)
      pdfMerge(
        [
          `./${outputDirectory}/${firstName}-${lastName}-overall.pdf`,
          `./${outputDirectory}/${firstName}-${lastName}-practicals.pdf`,
          `./${outputDirectory}/${firstName}-${lastName}-django-rest-react-opentdb-api.pdf`
        ],
        `./${outputDirectory}/${outputDirectory}-${firstName}-${lastName}-overall-results.pdf`
      )
        .then((_) =>
          console.log(`PDF files merged for ${firstName} ${lastName}.`.blue)
        )
        .catch((err) => console.log(err))
    }, idx * interval)
  })
}