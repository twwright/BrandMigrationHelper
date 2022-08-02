# Brand Migration Data Helper

# Purpose
When a client transitions from a single location to a multi-location, our team has to generate a new umbrella account called a Brand.
In order to move data from an individual location into a brand account, we have to export/import. However, the exported data 
needs to be cleaned, scrubbed, reconfgiured, and more.

This simple Flask app takes in the Location-specific .xlsx files then performs some actions to generate a Branch-specific .xlsx file
that can be immediately imported through our internal Brand Uploader. Prior to this tool, that work needed to be done manually.
