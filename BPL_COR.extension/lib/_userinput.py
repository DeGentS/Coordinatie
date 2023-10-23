import clr

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import Form, Label, TextBox, Button, SaveFileDialog, DialogResult, MessageBox
from System.Drawing import Size, Point
class InputForm(Form):
    def __init__(self):
        self.InitializeComponent()

    def InitializeComponent(self):
        self.Text = "User Input"
        self.Size = Size(700, 300)

        self.labels = []
        self.textBoxes = []

        # Maak een nieuw label voor de handleidingstekst
        self.guideLabel = Label()
        self.guideLabel.Text = "Vul de onderstaande velden in en klik op OK."
        self.guideLabel.Location = Point(350, 10)
        self.guideLabel.Size = Size(280, 40)  # Pas de grootte aan indien nodig
        self.Controls.Add(self.guideLabel)

        for i in range(3):
            label = Label()
            label.Text = "Enter Value {}:".format(i + 1)
            label.Location = Point(10, 20 + 50 * i)
            label.Size = Size(280, 20)
            self.Controls.Add(label)
            self.labels.append(label)

            textBox = TextBox()
            textBox.Location = Point(10, 40 + 50 * i)
            textBox.Size = Size(280, 20)
            self.Controls.Add(textBox)
            self.textBoxes.append(textBox)

        self.okButton = Button()
        self.okButton.Text = "OK"
        self.okButton.Location = Point(100, 170)
        self.okButton.Click += self.OkButtonClick
        self.Controls.Add(self.okButton)

    def SetLabelText(self, labelIndex, text):
        if 0 <= labelIndex < len(self.labels):
            self.labels[labelIndex].Text = text

    def SetFormTitle(self, title):
        self.Text = title

    def OkButtonClick(self, sender, args):
        self.values = [textBox.Text for textBox in self.textBoxes]
        self.DialogResult = DialogResult.OK

    def SetGuideText(self, text):
        self.guideLabel.Text = text

def get_user_input(title, label_texts, guide_text):
    form = InputForm()
    form.SetFormTitle(title)

    # Stel de handleidingstekst in
    form.SetGuideText(guide_text)

    # Pas de teksten van de labels aan
    for i, text in enumerate(label_texts):
        form.SetLabelText(i, text)

    if form.ShowDialog() == DialogResult.OK:
        return form.values
    else:
        return None, None, None