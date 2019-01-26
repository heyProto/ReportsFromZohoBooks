# ZOHO Books Integrations
Command line utility tool. Generates expense report for a particular project in Zoho Books. Built using [Node.js](https://nodejs.org/) v8.11.3

### Installation
```sh
$ npm install
```

### Usage 

```sh
$  node index.js --help
  Usage: index.js [options] [command]

  Commands:
    help     Display help
    version  Display version

  Options:
    -a, --attachment    Choose to download attachments for report (y or n)
    -e, --exchangerate  [Optional] USD rate to be used. Defaults to uncalculated if not mentioned.
    -h, --help          Output usage information
    -o, --orgid         Organization ID for Zoho Books account
    -p, --projectid     Project ID for generating report
    -t, --token         Access token to be used for pulling data
    -v, --version       Output the version number
```

###### Example
```sh
$ node index.js -t <zoho_access_token> -o <organization_id> -p <project_id> -a <y/n> -e <exchange_rate>
```