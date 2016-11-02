# Birthday Greeter

Sends out an email to birthday people in `Mitglieder.xls`. It requires the excel file to have the following fields:

 * Vorname
 * Nachname
 * Geburtsdatum (format `YYYY-MM-DD`)
 * E-Mail
 * Geschlecht (`== m` iff male)

## Usage

### Customizing

Edit the file `template.txt` to adjust the email sent. `$anrede` will be replaced by `Liebe $firstname` or `Lieber $firstname`, depending on the gender.

Generate a `Mitglieder.xls` (e.g. by intranet export) of the people you want to congratulate.

### Setup

The project is light on requirements, but we need something to parse excel files.

```
pip3 install -r requirements.txt
```


### Running the tests

Setting the environment variable `TEST` will run the tests.
```
TEST=1 python3 birthday.py
```

### Running

```
python3 birthday.py
```

This will send out an email to all birthday people. It will *not* keep track of who received a mail already, so make sure this only runs once a day.

In practice, you want to put this in `cron.daily`.


## Contributing

Pull requests are welcome :).
