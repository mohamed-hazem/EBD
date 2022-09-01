# Modules
# -- Main Modules -- #
from json import dumps, loads
from os import chdir, mkdir, path
from sys import argv, exit
from time import sleep

# -- PyAutoGui -- #
from pyautogui import click, size, typewrite, moveTo

# -- PyQt5 -- #
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication, QFileDialog, QMessageBox, QWidget
from PyQt5.QtCore import Qt

# -- Selenium -- #
from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from webdriver_manager.chrome import ChromeDriverManager

# -- GUI -- #
from ui import Ui_Form

# ---------------------------------------------------------------------------------------- #

# -- Main Variables -- #
SITE = "https://yale.egybest.fit/explore/"
quality_choose = {"1080p": "1", "720p": "2", "480p": "3", "360p": "4", "240p": "5"}
current_directory = path.dirname(__file__)
screen_width, screen_height = size()

# ---------------------------------------------------------------------------------------- #

chdir(current_directory) # Change current working directory

# -- Main Functions -- #
def validate_name(name):
    invalid_chars = [':', '*', '<', '>', '?', '"', '|']
    for char in invalid_chars:
        name = name.replace(char, '')
    return name

def make_directory(directory):
    directory = directory.replace("/", "\\")
    return directory

def get_range(input_range):
    input_range = input_range.split('-')
    start = int(input_range[0]) if input_range[0] != '' else 1
    end = int(input_range[-1]) if input_range[-1] != '' else None
    return start, end

def get_size_mb(size):
    size = size.replace(',', '')
    size = float(size[:-2]) * 1024 if size[-2] == 'G' else float(size[:-2])
    return size

def get_size(size):
    if size < 1024:
        size = round(size, 2)
        size = f"{size}MB"
    else:
        size = size / 1024
        size = round(size, 2)
        size = f"{size}GB"
    return size

def make_links_file(directory, name, size, links, season, e_start, e_end, e_r):
    size = get_size(sum(size))

    if e_r:
        if e_start == e_end:
            e_range = f"-E{str(e_start).zfill(2)}"
        else:
            e_range = f"-E{str(e_start).zfill(2)}-{str(e_end).zfill(2)}"
    else:
        e_range = ""

    file_name = f"{name}-S{str(season + 1).zfill(2)}{e_range}-{size}.txt"

    links_file = open(path.join(directory, file_name), 'a')
    for link in links:
        links_file.write(f"{link}\n")
    links_file.close()
# ------------------------------------- #

class MainBot:

    def __init__(self, search_key, quality, down_num, directory, IDM, IDM_extension_dir):
        
        # Attributes
        self.search_key = search_key
        self.quality = quality
        self.down_num = down_num
        self.directory = directory
        self.IDM = IDM
        self.IDM_extension_dir = IDM_extension_dir

        # Browser Options
        service = Service(ChromeDriverManager().install())
        options = Options()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.add_argument('--disable-notifications')
        options.add_extension(self.IDM_extension_dir) if self.IDM else None

        # Browser Capabilities
        caps = DesiredCapabilities().CHROME
        caps["pageLoadStrategy"] = "eager"

        # Open Browser
        self.browser = Chrome(service=service, options=options, desired_capabilities=caps)
        self.browser.maximize_window()
        self.browser.implicitly_wait(10)

    
    def main_step(self):
        search_query = f"{SITE}?q={self.search_key}"
        self.browser.get(search_query)

        results = self.browser.find_element(By.ID, "movies").find_elements(By.TAG_NAME, "a")
        results[self.down_num - 1].click()

        self.name = self.browser.find_element(By.CLASS_NAME, 'movie_title').get_attribute("innerText").title()

        self.close_ads()

    def series_step(self, season, episode, get_link=False):

        s_start, s_end = get_range(season)
        e_start, e_end = get_range(episode)

        seasons = self.browser.find_element(By.XPATH, '//*[@id="mainLoad"]/div[2]/div[2]/div').find_elements(By.TAG_NAME, 'a')
        seasons.reverse()
        s_end = len(seasons) if s_end == None else s_end

        for current_season in range(s_start - 1, s_end):
            self.browser.execute_script("arguments[0].click()", seasons[current_season])
            sleep(1)

            episodes = self.browser.find_element(By.CLASS_NAME, 'movies_small').find_elements(By.TAG_NAME, 'a')
            episodes.reverse()
            e_end = len(episodes) if e_end == None else e_end
            
            e_r = False if e_start == 1 and e_end == len(episodes) else True    
            self.download_links = []
            self.size = []

            for current_episode in range(e_start - 1, e_end):    
                episodes[current_episode].click()

                self.download(get_link=get_link)

                if  e_start < e_end:
                    episodes = self.browser.find_element(By.CLASS_NAME, 'movies_small').find_elements(By.TAG_NAME, 'a')
                    episodes.reverse()

            if get_link:
                make_links_file(self.directory, self.name, self.size, self.download_links, current_season, e_start, e_end, e_r)

                if (current_season == s_end - 1):
                    self.browser.quit()
                else:
                    e_start, e_end = 1, None
                    self.main_step()

                    seasons = self.browser.find_element(By.XPATH, '//*[@id="mainLoad"]/div[2]/div[2]/div').find_elements(By.TAG_NAME, 'a')
                    seasons.reverse()
    
    def download(self, get_link=False):
        sleep(2)

        if get_link:
            size_span = self.browser.find_element(By.XPATH, f'//*[@id="watch_dl"]/table/tbody/tr[{self.quality}]/td[3]').get_attribute("innerText")

           
        btn1 = self.browser.find_element(By.XPATH, f'//*[@id="watch_dl"]/table/tbody/tr[{self.quality}]/td[4]/a[1]')
        btn1.click()
        self.switch_to(1)

        btn2 = self.browser.find_element(By.XPATH, '/html/body/div[1]/div/p/a[1]')

        if (btn2.get_attribute('href') == None):
            btn2.click()
            self.switch_to(1)
            sleep(1)
            self.browser.refresh()

        btn3 = self.browser.find_element(By.XPATH, '/html/body/div[1]/div/p/a[1]')
        if not get_link:
            btn3.click()
        else:
            self.download_links.append(btn3.get_attribute('href'))
            self.size.append(get_size_mb(size_span))

            self.close_ads()

    def IDM_automation(self, dtype, season=None, episode=None):
            self.name = validate_name(self.name)

            if dtype == 'm':
                file_name = f"{self.name}.mp4"
                full_path = path.join(self.directory, self.name)
                abs_path = path.join(full_path, file_name)

                mkdir(full_path) if not path.isdir(full_path) else None
                    

            elif dtype == 's':
                file_name = f"Episode {episode}.mp4"
                series_dir = path.join(self.directory, self.name)
                season_dir = path.join(series_dir, f"Season {season}")
                abs_path = path.join(season_dir, file_name)

                mkdir(series_dir) if not path.isdir(series_dir) else None
                mkdir(season_dir) if not path.isdir(season_dir) else None
                    
            moveTo(screen_width*0.5, screen_height*0.5)
            sleep(0.5)
            click(screen_width*0.5, screen_height*0.45)
            sleep(0.5)
            typewrite(abs_path, interval=0.05)
            sleep(0.5)
            click(screen_width*0.5, screen_height*0.59)

            if self.IDM:
                self.browser.quit()
                
    # -- help functions -- #
    def switch_to(self, page):
        self.browser.switch_to.window(self.browser.window_handles[page])
    
    def close_ads(self):
        tabs = len(self.browser.window_handles)
        sleep(1)
        for _ in range(tabs - 1):
            self.switch_to(-1)
            self.browser.close()
        self.switch_to(0)

        


class Movie(MainBot):

    def __init__(self, search_key, quality, down_num, directory, IDM, IDM_extension_dir):

        super().__init__(search_key, quality, down_num, directory, IDM, IDM_extension_dir)

        self.main_step()
        self.download()
        self.IDM_automation('m')
        exit()

class Series(MainBot):

    def __init__(self, search_key, season, episode, quality, down_num, directory, IDM, IDM_extension_dir):

        self.season = season
        self.episode = episode
        super().__init__(search_key, quality, down_num, directory, IDM, IDM_extension_dir)

        self.main_step()
        self.series_step(self.season, self.episode)
        self.IDM_automation('s', self.season, self.episode)
        exit()

class AboNarer(MainBot):

    def __init__(self, search_key, seasons, episodes, quality, down_num, directory, IDM, IDM_extension_dir):

        self.seasons = seasons
        self.episodes = episodes
        super().__init__(search_key, quality, down_num, directory, IDM, IDM_extension_dir)

        self.main_step()
        self.series_step(self.seasons, self.episodes, get_link=True)
        exit()

# GUI class
class MainApp(QWidget, Ui_Form):
    def __init__(self, parent=None):
        super(MainApp, self).__init__(parent)
        QWidget.__init__(self)
        self.setupUi(self)

        self.ui()
        settings = self.get_default_settings()
        self.set_default_settings(settings)
        self.buttons()

        # Auto Focus on Startup
        self.MovieText.setFocus()
        
    # UI settings
    def ui(self):
        self.setFixedSize(823, 351)
        self.setWindowTitle("EgyBest Downloader")
        self.setWindowIcon(QIcon('icons/bot.png'))
        self.tabWidget.setTabIcon(1, QIcon('icons/settings.png'))
    
    # KeyPress Settings
    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Escape:
            self.close()
        
        # Download on press 'Enter'
        if e.key() in (Qt.Key_Return, Qt.Key_Enter):
            if self.tabWidget_2.currentIndex() == 0:
                self.movie()
            elif self.tabWidget_2.currentIndex() == 1:
                self.series()
            elif self.tabWidget_2.currentIndex() == 2:
                self.abo_narer()
        
        if e.key() == Qt.Key_Up:
            current_index = self.tabWidget_2.currentIndex()
            self.tabWidget_2.setCurrentIndex(current_index+1)
            self.focus_on()

        if e.key() == Qt.Key_Down:
            current_index = self.tabWidget_2.currentIndex()
            self.tabWidget_2.setCurrentIndex(current_index-1)
            self.focus_on()

    def focus_on(self):
        current_index = self.tabWidget_2.currentIndex()
        if current_index == 0:
            self.MovieText.setFocus()
        elif current_index == 1:
            self.SeriesText.setFocus()
        elif current_index == 2:
            self.SeriesScrapText.setFocus()

    # -------------------------- #

    # Browse Functions
    def browse(self, directory_input):
        directory = QFileDialog.getExistingDirectory(self, caption='Save To', directory='.')
        directory_input.setText(make_directory(directory))
    # -------------------------- #

    def IDM_browse(self):
        directory = QFileDialog.getOpenFileName(self, caption='Save To', directory='C:\Program Files (x86)\Internet Download Manager')
        self.IDMDDir.setText(make_directory(directory))
    # -------------------------- #

    # - Main Apps - # 
    def movie(self):

        movie = self.MovieText.text()
        quality = quality_choose[self.MovieQuality.currentText()]
        down_num = self.ResultNumber.value()
        directory = make_directory(self.MovieDir.text())
        IDM = self.IDMAutomation.isChecked()
        IDM_extension_dir = self.IDMDir.text()

        Movie(movie, quality, down_num, directory, IDM, IDM_extension_dir)
    
    def series(self):

        series = self.SeriesText.text()
        season = self.Season.text()
        episode = self.Episode.text()

        quality = quality_choose[self.SeriesQuality.currentText()]
        down_num = self.ResultNumber.value()
        directory = make_directory(self.SeriesDir.text())
        IDM = self.IDMAutomation.isChecked()
        IDM_extension_dir = self.IDMDir.text()

        Series(series, season, episode, quality, down_num, directory, IDM, IDM_extension_dir)

    def abo_narer(self):

        series = self.SeriesScrapText.text()
        seasons = self.ScrapSeasons.text()
        episodes = self.ScrapEpisodes.text()

        quality = quality_choose[self.SeriesScrapQuality.currentText()]
        down_num = self.ResultNumber.value()
        directory = make_directory(self.LinksDirSettings.text())
        
        AboNarer(series, seasons, episodes, quality, down_num, directory, None, None)
    # -------------------------- #

    # Settings Functions
    def get_default_settings(self):
        settings_file = open('settings.json')
        settings = loads(settings_file.read())
        settings_file.close()

        return settings

    def set_default_settings(self, settings):
        quality = settings['default_quality']
        movies_directory = settings['default_movies_directory']
        series_directory = settings['default_series_directory']
        links_directory = settings['default_links_directory']
        IDM_automation = settings['IDM_automation']
        IDM_extension_dir = settings['IDM_extension_dir']

        self.QualitySettings.setCurrentText(quality)
        self.MovieDirSettings.setText(movies_directory)
        self.SeriesDirSettings.setText(series_directory)
        self.LinksDirSettings.setText(links_directory)
        self.IDMAutomationSettings.setChecked(IDM_automation)

        self.MovieQuality.setCurrentText(quality)
        self.SeriesQuality.setCurrentText(quality)
        self.SeriesScrapQuality.setCurrentText(quality)

        self.MovieDir.setText(movies_directory)
        self.SeriesDir.setText(series_directory)
        self.IDMAutomation.setChecked(IDM_automation)
        self.IDMDir.setText(IDM_extension_dir)

    def set_settings(self):
        quality = self.QualitySettings.currentText()
        movies_directory = self.MovieDirSettings.text()
        series_directory = self.SeriesDirSettings.text()
        links_directory = self.LinksDirSettings.text()
        IDM_automation = self.IDMAutomationSettings.isChecked()
        IDM_extension_dir = self.IDMDir.text()

        self.MovieQuality.setCurrentText(quality)
        self.SeriesQuality.setCurrentText(quality)
        self.SeriesScrapQuality.setCurrentText(quality)

        self.MovieDir.setText(movies_directory)
        self.SeriesDir.setText(series_directory)
        self.IDMAutomation.setChecked(IDM_automation)

        settings = {
            "default_quality": quality,
            "default_movies_directory": movies_directory,
            "default_series_directory": series_directory,
            "default_links_directory": links_directory,
            "IDM_automation": IDM_automation,
            "IDM_extension_dir": IDM_extension_dir
        }

        settings = dumps(settings)

        settings_file = open('settings.json', "w")
        settings_file.write(settings)
        settings_file.close()

        QMessageBox.about(self, "Success !", "Settings has been updated")
    # -------------------------- #

    # Buttons Functions
    def buttons(self):
        # Browse Buttons
        self.MovieDirButton.clicked.connect(lambda: self.browse(self.MovieDir))
        self.SeriesDirButton.clicked.connect(lambda: self.browse(self.SeriesDir))
        self.IDMDirButton.clicked.connect(self.IDM_browse)
        self.MovieDirButtonSettings.clicked.connect(lambda: self.browse(self.MovieDirSettings))
        self.SeriesDirButtonSettings.clicked.connect(lambda: self.browse(self.SeriesDirSettings))
        self.LinksDirButtonSettings.clicked.connect(lambda: self.browse(self.LinksDirSettings))

        # Main Buttons
        self.MovieButton.clicked.connect(self.movie)
        self.SeriesButton.clicked.connect(self.series)
        self.ScraperButton.clicked.connect(self.abo_narer)

        # Apply Settings
        self.ApplySettings.clicked.connect(self.set_settings)

        # close buttons
        self.CloseButton.clicked.connect(lambda: self.close())
        self.CloseButton_2.clicked.connect(lambda: self.close())
    # -------------------------- #

def main():
    app = QApplication(argv)
    window = MainApp()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()
