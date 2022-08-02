# Brand Migration Data Helper

# Purpose
When a client transitions from a single live location to a multi-location, our team generates a new umbrella Brand account. In order to move data from an individual live locations into a single brand account, we have to export a raw xlsx file, then manipulate, clean, scrub, etc the file to create a new file that is properly prepared for import. 

This simple Flask app takes in the Location-specific .xlsx files from an export then performs some actions to generate a Brand-specific .xlsx file that can be immediately imported through our internal Brand Uploader. Prior to this tool, that work needed to be done manually over multiple hours.
