# Atzīmju iegūšana no ORTUS automatizācija
Šis projekts paredzēts, lai automatizētu atzīmju iegūšanu no RTU ORTUS platformas. Skripts pieslēdzas ORTUS platformai ar lietotājvārdu un paroli, meklē pašreizējā/pēdējā semestra kursus un iegūst atzīmes no katra kursa. Rezultāti tiek saglabāti Excel formātā failā atzimes.xlsx. Projekts arī dod iespēju daudz pārskatāmāk redzēt savas atzīmes vienā failā un nevajag manuāli navigēt mājaslapā un meklēt konkrētu atzīmi. Projekta ideja radās tāpēc, ka nebiju ilgu laiku universitātē un vajadzēja pārskatīt savus neizdarītos darbus/parādus. Atzīmju pārskatīšana ORTUS platformā ir šausmīgi neērta un radījās arī šis skripts.

## Prasības
* Python 3.x
* Google Chrome pārlūks
* Bibliotēkas
  * selenium
  * openpyxl
 
## Instalācija
1. Klonēt repositoriju
```
git clone https://github.com/traktoors/markAutomatization.git
```
2. Pāriet uz projekta direktoriju
```
cd markAutomatization
```
3. Instalējiet nepieciešamās bibliotēkas
```
pip install -r requirements.txt
```

## Lietošana
Palaist skriptu no projekta direktorija
```
python betterortus.py
```
Programma prasīs ievadīt ORTUS lietotājvārdu un paroli ( parole ir neredzama to ievadot! ), skripts automātiski pieslēgsies ORTUS platformai un
sāks meklēt atzīmes, ja lietotājvārds un parole būs korekti. Kad datu iegūšana ir veiksmīga un pabeigta, dati tiks saglabāti atzīmes.xlsx

## Kā strādā programma
Programma pēc lietotājvārda un paroles ievades uzsāks privātu interneta pārlūka sesiju ar selenium un ievadīs pieteikšanās datus. Pēc veiksmīgas pieteikšanās, selenium navigēs uz studentu lapu, kur ir norādīti kursi.
```
pieteikšanās lapa -> home lapa -> studentu lapa
```
Driveris atradīs pēdējo/šobrīdējo semestri un izies cauri katram kursam.
```
atver kursa linku -> navigē uz atzīmju linku
```
un šādi katram kursam.
ORTUS mājaslapā atzīmju vietnē atzīmes ir izkārtotas tabulas veidā, tāpēc ļoti vienkārši algoritmiski pārbauda un iegūst datus no katras līnijas. Dažas no līnijām tiek ignorētas, kuras tiek uzskatītas par tukšām - nav pietiekami daudz kolonnas < 3

Dati tiek paralēli saglabāti datu iegūšanas laikā un tiek saglabāti struktūrā
``` python
class Course:
    def __init__(self, courseName: str):
        self.courseName = courseName
        self.marks = []

    def addMark(self, title: str, value: str):
        self.marks.append((title, value))

    def getMarks(self):
        return self.marks

class CourseMap:
    def __init__(self):
        self.courses = []
    
    def addCourse(self, course: Course):
        self.courses.append(course)
    
    def getCourses(self):
        return self.courses
```
Katram kursa klasei ir lists ar testiem un atzīmēm, kas tiek saglabāts CourseMap objektā, kurš satur visus kursus un to iegūtos datus listā.
Kad visi dati tiek iegūti, sākās fāze saglabāšana excel failā, kura vienkārši iegūst no CourseMap visus kursu, un iterējot cauri katram kursam, to datus saglabā excel formātā.

## Funkcijas
* Automatizēta pieteikšanās ORTUS platformā
* Pašreizējā/pēdējā semestra kursu noteikšana
* Atzīmju iegūšana no katra kursa
* Atzīmju saglabāšana labākai pārskatīšanai Excel formātā

## Izmantotās bibliotēkas
* Selenium - Izmantots, lai iegūtu autentifikācijas tokenu, lai varētu tālāk iegūt savas atzīmes. Selenium ir ideāls risinājums, jo tas atveido lietotāja darbības.
* Openpyxl - Excel failu ģenerēšanai un formatēšanai
