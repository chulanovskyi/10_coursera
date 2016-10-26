# 10_coursera

###Prerequisites:

Run in console `pip install -r requirements.txt` to install 3rd party modules.

---

Script uses coursera xml feed to find some random courses and save info about them into Excel file.

###How to use:

You can edit in-script constants to get result as you want.

| Constant | Description |
| --- | --- |
| `COURSE_COUNT` | how many random courses need to find |

Run script in terminal like usual with `python coursera.py` and you will able to see progress:

```
Getting course list...
Done!
Getting 20 courses info...
Get: https://www.coursera.org/learn/thermo-apps
Get: https://www.coursera.org/learn/stem
Get: https://www.coursera.org/learn/negocios-internacionales-2
...
Done!
Making excel file...
Finish!
```

After that you can check newly created file `Coursera.xlsx` to view collected info about random courses.
