import win32com.client

def create_portal_frame():
    try:
        # Open Robot Structural Analysis
        robot = win32com.client.Dispatch("Robot.Application")
        robot.Visible = True
        
        # Create a new project
        project = robot.Project
        project.New()

        # Get user input for portal frame dimensions
        print("Enter dimensions for the portal frame:")
        span = float(input("Span (m): "))
        height = float(input("Height (m): "))
        column_section = input("Column section (e.g., HEA300): ")
        beam_section = input("Beam section (e.g., IPE300): ")
        material = input("Material (e.g., Steel): ")

        # Initialize model and define units (meters, kN)
        model = project.Structure
        model.Settings.Units.Set(3, 2, 2, 2)  # Length: m, Force: kN, Moment: kNm
        
        # Add nodes
        n1 = model.Nodes.Create(1, 0, 0, 0)  # Base of the left column
        n2 = model.Nodes.Create(2, 0, height, 0)  # Top of the left column
        n3 = model.Nodes.Create(3, span, height, 0)  # Top of the right column
        n4 = model.Nodes.Create(4, span, 0, 0)  # Base of the right column

        # Add bars (columns and beam)
        bars = model.Bars
        column1 = bars.Create(1, n1, n2)  # Left column
        beam = bars.Create(2, n2, n3)  # Beam
        column2 = bars.Create(3, n3, n4)  # Right column

        # Assign sections and materials
        sections = model.Sections
        sections.Create(column_section, material)
        sections.Create(beam_section, material)

        column1.SetSectionByName(column_section)
        beam.SetSectionByName(beam_section)
        column2.SetSectionByName(column_section)

        # Add boundary conditions (fixed supports)
        supports = model.Supports
        support = supports.Create(1)
        support.TranslationX = True
        support.TranslationY = True
        support.TranslationZ = True
        support.RotationX = True
        support.RotationY = True
        support.RotationZ = True
        model.Nodes.Get(1).Support = support
        model.Nodes.Get(4).Support = support

        # Update the model
        model.Update()

        print("Portal frame successfully created in Robot Structural Analysis!")
    except Exception as e:
        print(f"An error occurred: {e}")

# Run the function
create_portal_frame()
