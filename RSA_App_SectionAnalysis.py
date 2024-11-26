import win32com.client

# Corrected project types with simplified names for user prompt
project_types = {
    1: "FRAME_2D",
    2: "TRUSS_2D",
    3: "GRILLAGE",
    4: "FRAME_3D",
    5: "TRUSS_3D",
    6: "PLATE",
    7: "SHELL",
    8: "AXISYMMETRIC",
    9: "VOLUMETRIC",
    10: "CONCRETE_BEAM",
    11: "CONCRETE_COLUMN",
    12: "FOUNDATION",
    13: "PARAMETRIZED",
    14: "STEEL_CONNECTION",
    15: "SECTION",
    16: "PLANE_STRESS",
    17: "PLANE_DEFORMATION",
    18: "CONCRETE_DEEP_BEAM"
}

# Mapping simplified names back to the full formatted names for RSA
formatted_project_types = {key: f"I_PT_{value}" for key, value in project_types.items()}

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
                print(f"{formatted_project_types[user_choice]} ({project_types[user_choice]}) project created successfully!")
            else:
                print(f"An active {formatted_project_types[user_choice]} ({project_types[user_choice]}) project already exists.")
        except Exception as e:
            print(f"Failed to initialize Robot Structural Analysis or create the project: {e}")
    except ValueError:
        print("Invalid input. Please enter a valid project type number.")

# Call the function
initialize_robot_with_project_type()
