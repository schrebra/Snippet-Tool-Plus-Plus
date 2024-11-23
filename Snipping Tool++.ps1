Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create Snippets folder in Pictures if it doesn't exist
$snippetsFolder = Join-Path ([Environment]::GetFolderPath("MyPictures")) "Snippets"
if (-not (Test-Path $snippetsFolder)) {
    New-Item -ItemType Directory -Path $snippetsFolder | Out-Null
}

# Global variable for last screenshot path
$script:lastScreenshotPath = $null

# C# class for transparent form and selection
$csharpCode = @"
using System;
using System.Drawing;
using System.Windows.Forms;

public class SelectionForm : Form
{
    private Point selectStart;
    private Rectangle selection;
    private bool selecting = false;
    public event EventHandler<Rectangle> SelectionComplete;

    public SelectionForm()
    {
        this.TopMost = true;
        this.FormBorderStyle = FormBorderStyle.None;
        this.ShowInTaskbar = false;
        this.StartPosition = FormStartPosition.Manual;
        this.Location = new Point(0, 0);
        this.Size = Screen.PrimaryScreen.Bounds.Size;
        this.BackColor = Color.Black;
        this.Opacity = 0.4;
        this.Cursor = Cursors.Cross;
        this.DoubleBuffered = true;

        this.MouseDown += new MouseEventHandler(SelectionForm_MouseDown);
        this.MouseMove += new MouseEventHandler(SelectionForm_MouseMove);
        this.MouseUp += new MouseEventHandler(SelectionForm_MouseUp);
        this.Paint += new PaintEventHandler(SelectionForm_Paint);
    }

    private void SelectionForm_MouseDown(object sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            selecting = true;
            selectStart = e.Location;
            selection = new Rectangle();
        }
        else if (e.Button == MouseButtons.Right)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }

    private void SelectionForm_MouseMove(object sender, MouseEventArgs e)
    {
        if (selecting)
        {
            selection = new Rectangle(
                Math.Min(selectStart.X, e.X),
                Math.Min(selectStart.Y, e.Y),
                Math.Abs(e.X - selectStart.X),
                Math.Abs(e.Y - selectStart.Y));

            this.Invalidate();
        }
    }

    private void SelectionForm_MouseUp(object sender, MouseEventArgs e)
    {
        if (selecting && e.Button == MouseButtons.Left)
        {
            selecting = false;
            if (SelectionComplete != null)
            {
                SelectionComplete(this, selection);
            }
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }

    private void SelectionForm_Paint(object sender, PaintEventArgs e)
    {
        if (selecting)
        {
            using (Pen pen = new Pen(Color.Red, 4))
            {
                e.Graphics.DrawRectangle(pen, selection);
            }
        }
    }

    public Rectangle GetSelection()
    {
        return selection;
    }
}
"@

# Compile the C# code
Add-Type -TypeDefinition $csharpCode -ReferencedAssemblies System.Windows.Forms, System.Drawing

# Function to capture screenshot
function Capture-Screenshot {
    param (
        [System.Drawing.Rectangle]$bounds
    )

    $screenshot = New-Object System.Drawing.Bitmap($bounds.Width, $bounds.Height)
    $graphics = [System.Drawing.Graphics]::FromImage($screenshot)
    $graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
    $graphics.Dispose()
    return $screenshot
}

# Function to copy image to clipboard
function Copy-ImageToClipboard {
    param (
        [System.Drawing.Image]$image
    )

    if ($image -ne $null) {
        [System.Windows.Forms.Clipboard]::SetImage($image)
    }
}

# Function to add outline to image
function Add-ImageOutline {
    param (
        [System.Drawing.Image]$image,
        [System.Drawing.Color]$color,
        [int]$width
    )

    $newWidth = $image.Width + ($width * 2)
    $newHeight = $image.Height + ($width * 2)
    $newImage = New-Object System.Drawing.Bitmap($newWidth, $newHeight)
    $graphics = [System.Drawing.Graphics]::FromImage($newImage)

    # Fill background with outline color
    $graphics.Clear($color)

    # Draw original image in center
    $graphics.DrawImage($image, $width, $width, $image.Width, $image.Height)

    $graphics.Dispose()
    return $newImage
}

# Create the main form
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = " Snipping Tool++"
$mainForm.Size = New-Object System.Drawing.Size(550, 220)
$mainForm.StartPosition = "CenterScreen"
$mainForm.FormBorderStyle = "FixedDialog"
$mainForm.MaximizeBox = $false

# Button dimensions and spacing
$buttonWidth = 150
$buttonHeight = 40
$buttonSpacing = 25
$totalButtonsWidth = ($buttonWidth * 3) + ($buttonSpacing * 2)

# Calculate positions for centered buttons
$firstButtonX = ($mainForm.ClientSize.Width - $totalButtonsWidth) / 1.71
$secondButtonX = $firstButtonX + $buttonWidth + $buttonSpacing
$thirdButtonX = $secondButtonX + $buttonWidth + $buttonSpacing
$buttonY = 17

# Create the New Snip button
$newSnipButton = New-Object System.Windows.Forms.Button
$newSnipButton.Location = New-Object System.Drawing.Point([int]$firstButtonX, $buttonY)
$newSnipButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
$newSnipButton.Text = "New Snippet"
$mainForm.Controls.Add($newSnipButton)

# Create the Open with Photo Viewer button
$openButton = New-Object System.Windows.Forms.Button
$openButton.Location = New-Object System.Drawing.Point([int]$secondButtonX, $buttonY)
$openButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
$openButton.Text = "Open Last Snippet"
$openButton.Enabled = $false
$mainForm.Controls.Add($openButton)

# Create the Open Snippets Folder button
$openFolderButton = New-Object System.Windows.Forms.Button
$openFolderButton.Location = New-Object System.Drawing.Point([int]$thirdButtonX, $buttonY)
$openFolderButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
$openFolderButton.Text = "Open Snippets Folder"
$mainForm.Controls.Add($openFolderButton)

# Create outline options group
$outlineGroup = New-Object System.Windows.Forms.GroupBox
$outlineGroup.Text = "Outline Options"
$outlineGroup.Location = New-Object System.Drawing.Point(20, ($buttonY + $buttonHeight + 15))
$groupBoxWidth = $mainForm.ClientSize.Width - 32
$outlineGroup.Size = New-Object System.Drawing.Size($groupBoxWidth, 85)  # Increased height
$mainForm.Controls.Add($outlineGroup)

# Calculate vertical center position for controls
$controlsBaseY = 35  # Increased Y position for vertical centering

# Create outline checkbox
$outlineCheckbox = New-Object System.Windows.Forms.CheckBox
$outlineCheckbox.Text = "Add Outline"
$outlineCheckbox.Location = New-Object System.Drawing.Point(10, $controlsBaseY)
$outlineCheckbox.Size = New-Object System.Drawing.Size(100, 20)
$outlineCheckbox.Checked = $true
$outlineGroup.Controls.Add($outlineCheckbox)

# Create color button
$colorButton = New-Object System.Windows.Forms.Button
$colorButton.Text = "Select Color"
$colorButton.Location = New-Object System.Drawing.Point(120, $controlsBaseY)
$colorButton.Size = New-Object System.Drawing.Size(80, 23)
$outlineGroup.Controls.Add($colorButton)

# Create color preview panel
$colorPreview = New-Object System.Windows.Forms.Panel
$colorPreview.Location = New-Object System.Drawing.Point(210, $controlsBaseY)
$colorPreview.Size = New-Object System.Drawing.Size(23, 23)
$colorPreview.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$colorPreview.BackColor = [System.Drawing.Color]::Black
$outlineGroup.Controls.Add($colorPreview)

# Create width label
$widthLabel = New-Object System.Windows.Forms.Label
$widthLabel.Text = "Width:"
$widthLabel.Location = New-Object System.Drawing.Point(250, ($controlsBaseY + 3))
$widthLabel.Size = New-Object System.Drawing.Size(40, 20)
$outlineGroup.Controls.Add($widthLabel)

# Create width numeric updown
$widthUpDown = New-Object System.Windows.Forms.NumericUpDown
$widthUpDown.Location = New-Object System.Drawing.Point(290, $controlsBaseY)
$widthUpDown.Size = New-Object System.Drawing.Size(50, 23)
$widthUpDown.Minimum = 1
$widthUpDown.Maximum = 10
$widthUpDown.Value = 2
$outlineGroup.Controls.Add($widthUpDown)

# Calculate the vertical center position between outline group bottom and form bottom
$statusLabelY = $outlineGroup.Bottom + (($mainForm.ClientSize.Height - $outlineGroup.Bottom) / 2) - 10

# Create status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(0, $statusLabelY)
$statusLabel.Size = New-Object System.Drawing.Size($mainForm.ClientSize.Width, 30)
$statusLabel.Text = "Click 'New Screenshot' to start. Right-click to cancel."
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$mainForm.Controls.Add($statusLabel)

# Color button click event
$colorButton.Add_Click({
    $colorDialog = New-Object System.Windows.Forms.ColorDialog
    $colorDialog.Color = $colorPreview.BackColor
    if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $colorPreview.BackColor = $colorDialog.Color
    }
})

# Open button click event
$openButton.Add_Click({
    if ($script:lastScreenshotPath -and (Test-Path $script:lastScreenshotPath)) {
        Start-Process $script:lastScreenshotPath
    }
    else {
        $statusLabel.Text = "No screenshot available."
    }
})

# Open Folder button click event
$openFolderButton.Add_Click({
    if (Test-Path $snippetsFolder) {
        Start-Process $snippetsFolder
    }
    else {
        $statusLabel.Text = "Snippets folder not found."
    }
})

# New screenshot button click event
$newSnipButton.Add_Click({
    $mainForm.WindowState = "Minimized"
    Start-Sleep -Milliseconds 250

    $selectionForm = New-Object SelectionForm
    $result = $selectionForm.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $selection = $selectionForm.GetSelection()

        if ($selection.Width -gt 0 -and $selection.Height -gt 0) {
            $screenshot = Capture-Screenshot -bounds $selection

            # Add outline if checkbox is checked
            if ($outlineCheckbox.Checked) {
                $outlinedScreenshot = Add-ImageOutline -image $screenshot -color $colorPreview.BackColor -width $widthUpDown.Value
                $screenshot.Dispose()
                $screenshot = $outlinedScreenshot
            }

            # Generate filename with timestamp
            $timestamp = Get-Date -Format "yyMMdd_HHmmss"
            $filename = Join-Path $snippetsFolder "$timestamp.png"

            try {
                $screenshot.Save($filename, [System.Drawing.Imaging.ImageFormat]::Png)
                Copy-ImageToClipboard -image $screenshot
                $script:lastScreenshotPath = $filename
                $simpleFileName = "Pictures\Snippets\" + (Split-Path $filename -Leaf)
                $statusLabel.Text = "Saved to: $simpleFileName (Copied to clipboard)"
                $openButton.Enabled = $true
            }
            catch {
                $statusLabel.Text = "Error saving screenshot: $($_.Exception.Message)"
            }
            finally {
                $screenshot.Dispose()
            }
        }
        else {
            $statusLabel.Text = "Invalid selection area."
        }
    }
    else {
        $statusLabel.Text = "Screenshot cancelled."
    }

    $mainForm.WindowState = "Normal"
    $mainForm.Activate()
})

# Show the main form
$mainForm.ShowDialog()