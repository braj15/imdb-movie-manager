# imdb-movie-manager

A **Python** script for collecting and writing information `[Movie name, Imdb rating, Genre, Year, Actors, Director, Running time]` about movies (*in your local directory*) in an **Excel Table** where you can `Sort`, `Filter` etc. as you like and you can go straight **to the movie folder** from the excel sheet just by **clicking on the corresponding movie name**.

![xl-page](https://user-images.githubusercontent.com/37156545/41292784-3340ce4e-6e71-11e8-880f-c1a15cda11a0.png)

## Getting Started

Following instructions will help you get a copy of the project up and running on your machine.


### Prerequisites

* **Python 3** installed on your machine.  [[How to install Python 3](https://www.ics.uci.edu/~pattis/common/handouts/pythoneclipsejava/python.html)]


### Installation

* **Step 1** : Click the `Download ZIP` button from the `Clone or download` drop-down list as shown below :

![clone](https://user-images.githubusercontent.com/37156545/41351160-2ed350d0-6f33-11e8-8712-df32da54aa3b.png)

* **Step 2** : Unzip the downloaded file in your preferred location.

* **Step 3** : Install the **dependencies** :

  * Run the following commands on Command Prompt or you can install them in other way :
  
  ```
  pip install guessit
  pip install imdbpie
  pip install openpyxl
  pip install progressbar
  ```


### Usage


* **Step 1** : `Copy` the `collect_info.py` file to your local movie directory.

* **Step 2** : From the same directory, open Command Window using `Shift + Right click`.

* **Step 3** : Type `python collect_info.py` on command window and press `Enter`.        

* **Done** : Check for the `movie_info.xlsx` file in the same directory. Open it and use as you like.


   ![xlsheet](https://user-images.githubusercontent.com/37156545/41362355-cd5da262-6f4e-11e8-84ee-d4df1d500ff3.png)



## Issue

If the program gets stuck while running, check your internet connection and do Steps 2 & 3 of *Usage* section again.

## Author

* **Biraj Raj**

## License

This project is licensed under the MIT License - see the `LICENSE` file for details.

## Acknowledgments

Please feel free to suggest improvements. Happy Coding. :+1:
