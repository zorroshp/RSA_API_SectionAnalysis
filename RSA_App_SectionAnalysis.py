import win32com.client

# Corrected project types as per the document (pages 43–44)
project_types = {
    1: "I_PT_FRAME_2D",
    2: "I_PT_TRUSS_2D",
    3: "I_PT_GRILLAGE",
    4: "I_PT_FRAME_3D",
    5: "I_PT_TRUSS_3D",
    6: "I_PT_PLATE",
    7: "I_PT_SHELL",
    8: "I_PT_AXISYMMETRIC",
    9: "I_PT_VOLUMETRIC",
    10: "I_PT_CONCRETE_BEAM",
    11: "I_PT_CONCRETE_COLUMN",
    12: "I_PT_FOUNDATION",
    13: "I_PT_PARAMETRIZED",
    14: "I_PT_STEEL_CONNECTION",
    15: "I_PT_SECTION",
    16: "I_PT_PLANE_STRESS",
    17: "I_PT_PLANE_DEFORMATION",
    18: "I_PT_CONCRETE_DEEP_BEAM"
}

# Function to initialize RSA with the user-selected project type
def initialize_robot_with_project_type():
    # Display project types to the user
    print("Available Project Types:")
    for key, value in project_types.items():
        print(f"{key}: {value}")

    # Prompt user for project type
    try:
        user_choice = int(input("Enter the project type number to create: "))
        if user_choice not in project_types:
            print("Invalid project type selected. Exiting program.")
            return None

        # Initialize RSA application with the selected project type
        try:
            Robot = win32com.client.Dispatch("Robot.Application")
            Robot.Visible = True  # Make RSA visible
            Robot.Interactive = 1  # Enable interactive mode
            Robot.UserControl = True  # Allow user interaction
            print("Robot Structural Analysis initialized successfully!")

            # Check if a project is active or if it matches the selected type
            if not Robot.Project.IsActive or Robot.Project.Type != user_choice:
                Robot.Project.New(user_choice)
                print(f"{project_types[user_choice]} project created successfully!")
            else:
                print(f"An active {project_types[user_choice]} project already exists.")
        except Exception as e:
            print(f"Failed to initialize Robot Structural Analysis or create the project: {e}")
    except ValueError:
        print("Invalid input. Please enter a valid project type number.")

# Call the function
initialize_robot_with_project_type()
