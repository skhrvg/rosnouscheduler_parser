package main

import (
	"errors"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"os"
	"regexp"
	"strings"
	"time"
	"strconv"
	"encoding/json"

	"github.com/360EntSecGroup-Skylar/excelize"
)

type logger struct {
	debugLogger *log.Logger
	infoLogger  *log.Logger
	warnLogger  *log.Logger
	errorLogger *log.Logger
	fatalLogger *log.Logger
}

func (logger *logger) debug(s string, v ...interface{}) {
	logger.debugLogger.Printf(s, v...)
}
func (logger *logger) info(s string, v ...interface{}) {
	logger.infoLogger.Printf(s, v...)
}
func (logger *logger) warn(s string, v ...interface{}) {
	logger.warnLogger.Printf(s, v...)
}
func (logger *logger) error(s string, v ...interface{}) {
	logger.errorLogger.Printf(s, v...)
}
func (logger *logger) fatal(s string, v ...interface{}) {
	logger.fatalLogger.Printf(s, v...)
	os.Exit(1)
}

// Class - Занятие
type Class struct {
	Discipline string
	ClassType  string
	Date       time.Time
	Time       string
	Professor  string
	Subgroup   int
	Location   string
	Comment    string
}

// Group - Группа
type Group struct {
	GroupName         string
	NumberOfSubgroups int
	LastUpdate        time.Time
	Institute         string
	StudyLevel        string
	StudyForm         string
	Classes           []Class
}

func (group *Group) appendClass(class *Class) {
	//todo
}

// получить индекс первой строки расписания в таблице
func getFirstRowIndex(f *excelize.File) int {
	rows, _ := f.GetRows(f.GetSheetList()[0])
	var rowIndex int
	for rowIndex = 0; rowIndex < len(rows); rowIndex++ {
		if rows[rowIndex][1] == "ПОНЕДЕЛЬНИК" {
			break
		}
	}
	return rowIndex
}

// получить индекс последней строки расписания в таблице
func getLastRowIndex(f *excelize.File) (lastRowIndex int, err error) {
	rows, _ := f.GetRows(f.GetSheetList()[0])
	var emptyRowCounter, rowIndex int
	for rowIndex = getFirstRowIndex(f); rowIndex < len(rows) && emptyRowCounter < 6; rowIndex++ {
		if rows[rowIndex][0] == "" {
			emptyRowCounter++
		} else if rows[rowIndex][1] != "" {
			emptyRowCounter = 0
		} else {
			emptyRowCounter++
		}
	}
	if rows[rowIndex-4][1] == "СУББОТА" {
		lastRowIndex = rowIndex - 3
	} else {
		lastRowIndex = rowIndex - 5
	}
	return
}

// получить индекс последней колонки расписания в таблице
func getLastColIndex(f *excelize.File) int {
	cols, _ := f.GetCols(f.GetSheetList()[0])
	firstRowIndex := getFirstRowIndex(f)
	var emptyColCounter, colIndex int
	for colIndex = 2; colIndex < len(cols)-1 && emptyColCounter < 3; colIndex++ {
		if cols[colIndex][firstRowIndex] == "" {
			emptyColCounter++
		} else {
			emptyColCounter = 0
		}
	}
	return colIndex - emptyColCounter
}

// получить дату первой ячейки в расписании
func getFirstDateAndCol(f *excelize.File) (firstDate time.Time, firstCol int, err error) {
	cols, _ := f.GetCols(f.GetSheetList()[0])
	firstRowIndex := getFirstRowIndex(f)
	var colIndex, day int
	var month time.Month
	switch cols[2][firstRowIndex-3] {
	case "ЯНВАРЬ":
		month = time.January
	case "ФЕВРАЛЬ":
		month = time.February
	case "МАРТ":
		month = time.March
	case "АПРЕЛЬ":
		month = time.April
	case "МАЙ":
		month = time.May
	case "ИЮНЬ":
		month = time.June
	case "ИЮЛЬ":
		month = time.July
	case "АВГУСТ":
		month = time.August
	case "СЕНТЯБРЬ":
		month = time.September
	case "ОКТЯБРЬ":
		month = time.October
	case "НОЯБРЬ":
		month = time.November
	case "ДЕКАБРЬ":
		month = time.December
	default:
		err = errors.New("неизвестный первый месяц")
		return
	}
	for colIndex = 2; colIndex < 5; colIndex++ {
		if cols[colIndex][firstRowIndex] != "" {
			day, err = strconv.Atoi(cols[colIndex][firstRowIndex])
			firstDate = time.Date(time.Now().Year(), month, day, 0, 0, 0, 0, time.UTC)
			firstCol = colIndex
			return
		}
	}
	err = errors.New("нет числа первого месяца")
	return
}

// получить 3 массива: с индексами строк начала / окончания для каждого дня недели и с количеством возможных предметов для каждого дня недели
func getWeekdaysAndDisciplines(f *excelize.File) (weekdaysStart [6]int, weekdaysEnd [6]int, weekdaysDisciplines [6]int, err error) {
	cols, _ := f.GetCols(f.GetSheetList()[0])
	lastRowIndex, err := getLastRowIndex(f)
	if err != nil {
		return
	}
	// разметка по дням недели
	for rowIndex := getFirstRowIndex(f); rowIndex <= lastRowIndex; rowIndex++ {
		switch cols[1][rowIndex] {
		case "ПОНЕДЕЛЬНИК":
			weekdaysStart[0] = rowIndex + 1
		case "ВТОРНИК":
			weekdaysEnd[0] = rowIndex - 1
			weekdaysDisciplines[0] = (weekdaysEnd[0] - weekdaysStart[0] + 1) / 3
			weekdaysStart[1] = rowIndex + 1
		case "СРЕДА":
			weekdaysEnd[1] = rowIndex - 1
			weekdaysDisciplines[1] = (weekdaysEnd[1] - weekdaysStart[1] + 1) / 3
			weekdaysStart[2] = rowIndex + 1
		case "ЧЕТВЕРГ":
			weekdaysEnd[2] = rowIndex - 1
			weekdaysDisciplines[2] = (weekdaysEnd[2] - weekdaysStart[2] + 1) / 3
			weekdaysStart[3] = rowIndex + 1
		case "ПЯТНИЦА":
			weekdaysEnd[3] = rowIndex - 1
			weekdaysDisciplines[3] = (weekdaysEnd[3] - weekdaysStart[3] + 1) / 3
			weekdaysStart[4] = rowIndex + 1
		case "СУББОТА":
			weekdaysEnd[4] = rowIndex - 1
			weekdaysDisciplines[4] = (weekdaysEnd[4] - weekdaysStart[4] + 1) / 3
			weekdaysStart[5] = rowIndex + 1
			weekdaysEnd[5], _ = getLastRowIndex(f)
			weekdaysDisciplines[5] = (weekdaysEnd[5] - weekdaysStart[5] + 1) / 3
			break
		default:
			continue
		}
	}
	// проверка на ошибки
	for _, n := range weekdaysStart {
		if n == 0 {
			err = errors.New("ошибка при подсчете начальных строк дня недели")
			return
		}
	}
	for _, n := range weekdaysEnd {
		if n == 0 {
			err = errors.New("ошибка при подсчете последних строк дня недели")
			return
		}
	}
	for i := 0; i <= 5; i++ {
		if (weekdaysEnd[i]-weekdaysStart[i]+1)%3 != 0 && weekdaysEnd[i] != weekdaysStart[i] {
			err = errors.New("ошибка при подсчете количества предметов - возможно, в таблице некорректное количество строк")
			return
		}
	}
	return
}

// парсинг одного файла
func parseFile(file os.FileInfo, filePath string) (classes []Class, groupsInFile []string, err error) {
	fileName := file.Name()
	l.info("[%s] Начало парсинга таблицы...\n", fileName)
	// названия групп по названию файла
	groupsInFile = strings.Split(strings.Replace(strings.Replace(strings.Replace(strings.Replace(fileName, ".xlsx", "", -1), ", ", ",", -1), "б,п", "бп", -1), "п,б", "пб", -1), ",")
	l.info("[%s] Найдены группы: %s\n", fileName, groupsInFile)
	for _, group := range groupsInFile {
		if isMatchRegexp, _ := regexp.MatchString(`^(\d{3})([а-яА-Я]{0,3})(-\d|\d)?$`, group); !isMatchRegexp {
			l.warn("[%s] Некорректное название группы: \"%s\"!\n", fileName, group)
			err = errors.New("некорректное название группы в имени файла")
			return
		}
	}

	// получение разметки файла
	excelFile, err := excelize.OpenFile(filePath + "/" + fileName)
	if err != nil {
		l.warn("[%s] Не удалось открыть Excel файл: %s\n", fileName, err)
		err = errors.New("excelize не смог открыть файл")
		return
	}
	firstRowIndex := getFirstRowIndex(excelFile)
	lastRowIndex, err := getLastRowIndex(excelFile)
	firstDate, firstCol, err := getFirstDateAndCol(excelFile)
	if err != nil {
		l.warn("[%s] Ошибка при расчете разметки файла: %s\n", fileName, err)
		return
	}
	lastColIndex := getLastColIndex(excelFile)
	weekdaysStart, weekdaysEnd, weekdaysDisciplines, err := getWeekdaysAndDisciplines(excelFile)
	l.debug("FRI:%d LRI:%d LCI:%d WS:%d WE:%d WD:%d", firstRowIndex, lastRowIndex, lastColIndex, weekdaysStart, weekdaysEnd, weekdaysDisciplines)
	if err != nil {
		l.warn("[%s] Ошибка при расчете разметки файла: %s\n", fileName, err)
		return
	}
	
	// парсинг занятий
	cols, _ := excelFile.GetCols(excelFile.GetSheetList()[0])
	for weekday, disciplines := range weekdaysDisciplines {
		if disciplines == 0 {
			// постоянный выходной
			continue
		}
		for discipline := 0; discipline < disciplines; discipline++ {
			for col := firstCol; col <= lastColIndex; col++ {
				if cols[col][weekdaysStart[weekday] + discipline*3] != "" {
					currentClass := Class{
						Discipline: cols[1][weekdaysStart[weekday] + discipline*3],
						Time: cols[0][weekdaysStart[weekday] + discipline*3],
						ClassType: strings.Replace(cols[col][weekdaysStart[weekday] + discipline*3], "\n", " ", -1),
						Comment: cols[col][weekdaysStart[weekday] + discipline*3 + 1],
						Location: cols[1][weekdaysStart[weekday] + discipline*3 + 2],
						Professor: cols[1][weekdaysStart[weekday] + discipline*3 + 1],
						Date: firstDate.AddDate(0, 0, ((col-firstCol)*7)+weekday),
					}
					for strings.Contains(currentClass.ClassType, "  ") {
						currentClass.ClassType = strings.Replace(currentClass.ClassType, "  ", " ", -1)
					}
					switch currentClass.ClassType {
					case "Л":
						currentClass.ClassType = "Лекция"
					case "С":
						currentClass.ClassType = "Семинар"
					case "ПЗ":
						currentClass.ClassType = "Практическое занятие"
					case "ЗАЧ":
						currentClass.ClassType = "ЗАЧЕТ"
					case "Л/ПЗ":
						currentClass.ClassType = "Лекция / Практическое занятие"
					case "Л/С":
						currentClass.ClassType = "Лекция / Семинар"
					case "Лаб":
						currentClass.ClassType = "ЛАБОРАТОРНАЯ РАБОТА"
					case "ЛАБ":
						currentClass.ClassType = "ЛАБОРАТОРНАЯ РАБОТА"
					case "ДИФ.ЗАЧ":
						currentClass.ClassType = "ДИФ. ЗАЧЕТ"
					case "ЗАЩ":
						currentClass.ClassType = "ЗАЩИТА"
					case "С/Л":
						currentClass.ClassType = "Семинар / Лекция"
					case "ПЗ/Л":
						currentClass.ClassType = "Практическое занятие / Лекция"
					case "Л/ЗАЧ":
						currentClass.ClassType = "Лекция / ЗАЧЕТ"
					case "К":
						currentClass.ClassType = "КОНСУЛЬТАЦИЯ"
					case "ЭКЗ":
						currentClass.ClassType = "ЭКЗАМЕН"
					default:
						l.warn("[%s] Неизвестный тип занятия: %s [%d:%d]", fileName, currentClass.ClassType, col, weekdaysStart[weekday] + discipline*3)
					}
					classes = append(classes, currentClass)
				}
			}
		}
	}
	return
}

// парсинг всех файлов в директории ./cache/downloads
func parseDownloads() (groups []Group, err error) {
	l.info("Парсинг загруженных файлов...\n")
	timestamp := time.Now().Format("2006-01-02-15-04-05")
	filePath := "./cache/downloads"
	err = os.MkdirAll(filePath+"/../defective/"+timestamp+"/", 0755)
	if err != nil {
		l.error("Ошибка при создании директории: %s\n", err)
		return
	}
	err = os.MkdirAll(filePath+"/../parsed/"+timestamp+"/", 0755)
	if err != nil {
		l.error("Ошибка при создании директории: %s\n", err)
		return
	}
	files, err := ioutil.ReadDir(filePath)
	if err != nil {
		l.error("Ошибка при получении списка файлов: %s\n", err)
		return
	}
	for _, file := range files {
		if !file.IsDir() && strings.Contains(file.Name(), ".xlsx") {
			classes, groupsInFile, errSkip := parseFile(file, filePath)
			if errSkip != nil {
				l.warn("Пропуск таблицы \"%s\": %s.\n", file.Name(), errSkip)
				err = os.Rename(filePath+"/"+file.Name(), filePath+"/../defective/"+timestamp+"/"+file.Name())
				if err != nil {
					l.error("Ошибка при перемещении таблицы \"%s\". Парсинг остановлен.", file.Name())
					l.error(fmt.Sprint(err))
					return
				}
				continue
			}
			for _, group := range groupsInFile {
				var institute, studyLevel string
				switch string(group[0]) {
				case "1":
					institute = "Институт экономики, управления и финансов"
				case "2":
					institute = "Юридический институт"
				case "3":
					institute = "Институт бизнес-технологий"
				case "4":
					institute = "Институт информационных систем и инженерно-компьютерных технологий"
				case "5":
					institute = "Институт психологии и педагогики"
				case "6":
					institute = "Институт гуманитарных технологий"
				default:
					l.error("Неизвесный номер института: %s", file.Name())
					return
				}
				if strings.Contains(group, "м") {
					studyLevel = "Магистратура"
				} else {
					studyLevel = "Бакалавриат"
				}
				currentGroup := Group{
					GroupName: group,
					LastUpdate: time.Now(),
					Classes: classes,
					Institute: institute,
					NumberOfSubgroups: 0,
					StudyForm: "Очная",
					StudyLevel: studyLevel,
				}
				groups = append(groups, currentGroup)
			}
			err = os.Rename(filePath+"/"+file.Name(), filePath+"/../parsed/"+timestamp+"/"+file.Name())
			if err != nil {
				l.error("Ошибка при перемещении таблицы \"%s\". Парсинг остановлен.", file.Name())
				l.error(fmt.Sprint(err))
				return
			}
			l.info("Успешный парсинг таблицы \"%s\".\n", file.Name())
		}
	}
	l.info("Парсинг загруженных файлов окончен.\n")
	return
}

var l logger

func main() {
	// создание логгера
	logFile, err := os.OpenFile("log.txt", os.O_CREATE|os.O_APPEND|os.O_RDWR, 0666)
	if err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
	logMultiwriter := io.MultiWriter(os.Stdout, logFile)
	l = logger{
		debugLogger: log.New(logMultiwriter, "[DEBUG] ", log.Ldate|log.Ltime),
		infoLogger:  log.New(logMultiwriter, "[INFO]  ", log.Ldate|log.Ltime),
		warnLogger:  log.New(logMultiwriter, "[WARN]  ", log.Ldate|log.Ltime),
		errorLogger: log.New(logMultiwriter, "[ERROR] ", log.Ldate|log.Ltime),
		fatalLogger: log.New(logMultiwriter, "[FATAL] ", log.Ldate|log.Ltime),
	}
	// создание директорий
	err = os.MkdirAll("./cache/downloads", 0755)
	if err != nil {
		l.fatal("Ошибка при создании директории: %s", err)
	}

	groups, err := parseDownloads()
	if err != nil {
		l.fatal("Ошибка при парсинге загрузок: %s", err)
	}

	// тут будет отправка данных в БД
	tmp, _ := json.Marshal(groups)
	l.debug("%s", tmp)
}
