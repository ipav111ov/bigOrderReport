function onOpen() {
  SpreadsheetApp.getUi().createMenu("BigOrderReport")
    .addItem("Calculate", "main")
    .addToUi();
}

const BIGORDER_REQS = {
  cameras: 75,
  square: 500,
}

const COLUMNS = {
  indexDate: 0,
  indexOrderId: 1,
  indexPlatform: 2,
  indexCreator: 3,
  indexRecipients: 4,
  indexType: 5,
  indexMark: 6,
  indexComment: 7,
  indexSquare: 8,
  indexCameras: 9,
  indexSpentTime: 10,
  indexReviewSpentTime: 11,
  indexCreatorUid: 12,
  indexRecipientsUid: 13,
}

const TYPE_ARRAY = ['solo', 'shared', 'total']


const TRIM_PERCENT = 5
const FRACTION_DIGITS = 2

function main() {
  const bigOrderReport = new Data(getValuesFromSS())
  bigOrderReport.getCalculations()
  const sortedBigOrderReport = getSortedObjectByMonths(bigOrderReport)
  const reportForOutput = output(sortedBigOrderReport)
  return reportForOutput
}

class Data {
  constructor(values) {
    for (const orderAsRow of values) {

      const square = typeof orderAsRow[COLUMNS.indexSquare] == 'number' ? orderAsRow[COLUMNS.indexSquare] : typeof orderAsRow[COLUMNS.indexSquare] == 'string' ? Number(orderAsRow[COLUMNS.indexSquare]) : 0

      const cameras = typeof orderAsRow[COLUMNS.indexCameras] == 'number' ? orderAsRow[COLUMNS.indexCameras] : typeof orderAsRow[COLUMNS.indexCameras] == 'string' ? Number(orderAsRow[COLUMNS.indexCameras]) : 0

      if (square > BIGORDER_REQS.square || cameras > BIGORDER_REQS.cameras) {

        const date = moment(orderAsRow[COLUMNS.indexDate])

        const monthYear = date.format("MMMM YYYY")

        const spentTime = typeof orderAsRow[COLUMNS.indexSpentTime] == 'number' ? orderAsRow[COLUMNS.indexSpentTime] : typeof orderAsRow[COLUMNS.indexSpentTime] == 'string' ? Number(orderAsRow[COLUMNS.indexSpentTime]) : 0

        const orderType = typeof orderAsRow[COLUMNS.indexType] == 'string' ? orderAsRow[COLUMNS.indexType] : 'orderType is not defined'

        const recipientsUid = orderAsRow[COLUMNS.indexRecipientsUid]

        const recipientsUidArray = recipientsUid ? recipientsUid.split(',') : []

        if (!this[orderType]) {
          this[orderType] = {}
        }
        if (!this[orderType][monthYear]) {
          const month = new Month()
          this[orderType][monthYear] = month
        }

        if (recipientsUidArray) {
          const type = recipientsUidArray.length == 1 ? 'solo' : 'shared'
          this[orderType][monthYear][type].time += spentTime
          this[orderType][monthYear][type].arrayTime.push(spentTime)
          this[orderType][monthYear][type].cameras += cameras

        }
      }
    }
  }



  getCalculations() {
    for (const orderType of Object.keys(this)) {
      for (const month of Object.keys(this[orderType])) {
        for (const type of TYPE_ARRAY) {

          if (type == "total") {
            this[orderType][month][type].time = this[orderType][month].solo.time + this[orderType][month].shared.time

            this[orderType][month][type].arrayTime = [...this[orderType][month].solo.arrayTime, ...this[orderType][month].shared.arrayTime]

            this[orderType][month][type].cameras = this[orderType][month].solo.cameras + this[orderType][month].shared.cameras
          }

          this[orderType][month][type].median = Number(getMedian(this[orderType][month][type].arrayTime).toFixed(FRACTION_DIGITS))

          this[orderType][month][type].average = Number(getAverage(this[orderType][month][type].arrayTime).toFixed(FRACTION_DIGITS))

          this[orderType][month][type].trimmedAverage = Number(getTrimmedAverage(this[orderType][month][type].arrayTime, TRIM_PERCENT).toFixed(FRACTION_DIGITS))

          this[orderType][month][type].speed = Number(getSpeed(this[orderType][month][type].time, this[orderType][month][type].cameras).toFixed(FRACTION_DIGITS))

        }

      }
    }
  }
}


class Month {
  constructor() {
    this.solo = new Type();
    this.shared = new Type();
    this.total = new Type();

  }
}

class Type {
  constructor() {
    this.median = 0;
    this.average = 0;
    this.trimmedAverage = 0;
    this.time = 0;
    this.cameras = 0;
    this.arrayTime = [];
  }
}

function output(report) {
  const ss = SpreadsheetApp.getActive();

  for (const orderType of Object.keys(report)) {

    if (!ss.getSheetByName(orderType)) {
      ss.insertSheet(orderType)
    }
    const sheet = ss.getSheetByName(orderType)

    arrayForWrite = [
      [orderType, 'Solo', '', '', '', '', '', '', 'Shared', '', '', '', '', '', '', 'Total', '', '', '', '', '', ''],
      ['Month', 'Median', 'Average', 'Trimmed', 'Time Solo', 'Cameras Solo', 'Speed Solo', 'Orders Solo', 'Median', 'Average', 'Trimmed', 'Time Shared', 'Cameras Shared', 'Speed Shared', 'Orders Solo', 'Median', 'Average', 'Trimmed', 'Time', 'Cameras', 'Speed', 'Orders']
    ];

    for (const month of Object.keys(report[orderType])) {

      let arr = [month]
      for (type of TYPE_ARRAY) {
        const median = report[orderType][month][type].median
        const average = report[orderType][month][type].average
        const trimmedAverage = report[orderType][month][type].trimmedAverage
        const time = report[orderType][month][type].time
        const cameras = report[orderType][month][type].cameras
        const speed = report[orderType][month][type].speed
        const orders = report[orderType][month][type].arrayTime.length

        arr.push(median, average, trimmedAverage, time, cameras, speed, orders)
      }
      arrayForWrite.push(arr)
    }

    sheet.getDataRange().clear()
    sheet.getRange(1, 1, arrayForWrite.length, arrayForWrite[0].length).setValues(arrayForWrite)

  }
  return Logger.log('Completed')
}

function getSortedObjectByMonths(object) {
  let sortedObject = {}

  for (const orderType of Object.keys(object)) {
    let keys = Object.keys(object[orderType])
    sortedObject[orderType] = {}

    keys.sort((a, b) => new Date(a) - new Date(b))
    keys.forEach(month => {
      sortedObject[orderType][month] = object[orderType][month]
    })
  }
  return sortedObject
}

function getMedian(array) {
  if (array.length === 0) return 0
  const copyArray = [...array]
  const half = Math.floor(copyArray.length / 2)
  return copyArray.sort((a, b) => a - b).length % 2 == 1 ? copyArray[half] : (copyArray[half] + copyArray[half - 1]) / 2;
}

function getAverage(array) {
  if (array.length === 0) return 0
  const sum = array.reduce((acc, current) => acc += current)
  return sum / array.length
}

function getTrimmedAverage(array, trimPercent) {
  if (array.length === 0) return 0
  const copyArray = [...array]
  const trimCount = Math.floor((trimPercent / 2 * 0.01) * copyArray.length);
  if (!trimCount) return getAverage(copyArray)
  return getAverage(copyArray.sort((a, b) => a - b).slice(trimCount, copyArray.length - trimCount))
}

function getSpeed(time, cameras) {
  if (time !== 0 && cameras !== 0) {
    return time / cameras
  }
  return 0
}