import win32com.client

oAnsoftApp = win32com.client.Dispatch("AnsoftMaxwell.MaxwellScriptInterface")
oDesktop = oAnsoftApp.GetAppDesktop()
oProject = oDesktop.NewProject()
oDesign = oProject.InsertDesign("Maxwell 3D", "ScriptTest", "Magnetostatic", "")
oEditor = oDesign.SetActiveEditor("3D Modeler")

# create the first cone
firstConeName = "firstCone"
coneBotRad = "1.5mm"
oEditor.CreateCone(
    {
        "NAME:ConeParameters": "",
        "XCenter": "0mm",
        "YCenter": "0mm",
        "ZCenter": "0mm",
        "WhichAxis": "Z",
        "Height": "2mm",
        "BottomRadius": coneBotRad,
        "TopRadius": "0mm"
    },
    {
        "NAME:Attributes": "",
        "Name": firstConeName,
        "Flags": "",
        "Color": "(132 132 193)",
        "Transparency": 0,
        "PartCoordinateSystem": "Global",
        "UDMId": "",
        "MaterialValue": "\"vacuum\"",
        "SolveInside": True
    }
)

# Now replicate this a few times and create an array out of it
for x in range(5):
    for y in range(5):
        # leave the first one alone in it's created
        # position
        if x == 0 and y == 0:
            continue
        # all other grid positions, replicate from the
        # first one
        # copy first
        oEditor.Copy(
            {
                "NAME": "Selections",
                "Selections": firstConeName
            }
        )
        # paste it and capture the pasted name
        # the pasted names come in an array as we could
        # be pasting a selection cmposed of multiple objects
        pasteName = oEditor.Paste()[0]
        # now move the pasted item to it's final position
        oEditor.Move(
            {"NAME": "Selections", "Selections": pasteName},
            {
                "NAME": "TransalateParameters",
                "CoordinateSystemID": -1,
                "TranslateVectorX": "%d * 3 * %s" % (x, coneBotRad),
                "TranslateVectorY": "%d * 3 * %s" % (y, coneBotRad),
                "TranslateVectorZ": "0mm"
            }
        )
        # Now fit the display to the created grid
oEditor.FitAll()
