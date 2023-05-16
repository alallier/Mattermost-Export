# Mattermost-Export

A simple Python script that exports Mattermost chat data to various formats via the Mattermost API

## Usage:

MMExport2PDF.py [options]

MMExport2PDF.py is used to export all a users channels and DMs from a team.

```
options:
  -h, --help            show this help message and exit

User Info:
  -a AUTH, --auth AUTH  Auth Token (default: None)
  -u USER, --user USER  Username of user to be exported (default: None)
  -t TEAM, --team TEAM  Team to export from (default: None)

Server Info:
  -s SERVER, --server SERVER
                        Hostname or IP of the server (default: mattermost.com)

Channel Categories:
  -p, --public          Exclude public channels
  -P, --private         Exclude private channels
  -g, --groups          Exclude group messages
  -d, --DMs             Exclude direct messages

Message Filters:
  -I [INCLUDE ...], --include [INCLUDE ...]
                        Only inlcude these channels in the export. (default:
                        [])
  -E [EXCLUDE ...], --exclude [EXCLUDE ...]
                        Exclude these channels from the export (default: [])

Export Options:
  -i, --images          Embed images in PDF (default: False)
  -f, --files           Embed files in PDF (default: False)
  -j, --json            Export JSON (default: False)
  -o OUTPUT, --output OUTPUT
                        Base output directory (default: ./users)
```

This can take a long time to run.
