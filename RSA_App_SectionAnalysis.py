import win32com.client

# Initialize the Robot Structural Analysis application
try:
    Robot = win32com.client.Dispatch("Robot.Application")
    Robot.Visible = True  # Make RSA visible
    Robot.Interactive = 1  # Enable interactive mode
    Robot.UserControl = True  # Allow user interaction
    print("Robot Structural Analysis initialized successfully!")
except Exception as e:
    print(f"Failed to initialize Robot Structural Analysis: {e}")
    exit()

# Check if a project is active and if it's not a shell project
try:
    SHELL_PROJECT_TYPE = 5  # Shell Design project type
    if not Robot.Project.IsActive or Robot.Project.Type != SHELL_PROJECT_TYPE:
        Robot.Project.New(SHELL_PROJECT_TYPE)  # Create a new Shell project
        print("Shell project created successfully!")
    else:
        print("An active shell project already exists.")
except Exception as e:
    print(f"Failed to create or initialize the project: {e}")
