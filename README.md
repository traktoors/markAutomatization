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
Programma prasīs ievadīt ORTUS lietotājvārdu un paroli, skripts automātiski pieslēgsies ORTUS platformai un
sāks meklēt atzīmes. Kad datu iegūšana ir veiksmīga un pabeigta, dati tiks saglabāti atzīmes.xlsx

## Funkcijas
* Automatizēta pieteikšanās ORTUS platformā
* Pašreizējā/pēdējā semestra kursu noteikšana
* Atzīmju iegūšana no katra kursa
* Atzīmju saglabāšana labākai pārskatīšanai Excel formātā

## Izmantotās bibliotēkas
* Selenium - Izmantots tieši šis, lai autentificētos ORTUS platformā un iegūtu nepieciešamos datus, jo ORTUS piekļuve ir aizsargāta ar autentifikāciju
* Openpyxl - Excel failu ģenerēšanai un formatēšanai
