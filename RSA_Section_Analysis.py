import win32com.client

def initialize_robot_frame_2d():
    try:
        Robot = win32com.client.Dispatch("Robot.Application")
        Robot.Visible = True
        Robot.Interactive = 1
        Robot.UserControl = True
        print("Robot Structural Analysis initialized successfully!")
        if not Robot.Project.IsActive or Robot.Project.Type != 1:
            Robot.Project.New(1)
            print("Frame 2D project created successfully!")
        return Robot
    except Exception as e:
        print(f"Failed to initialize Robot Structural Analysis: {e}")
        exit()

def get_input(prompt, default_value):
    user_input = input(prompt)
    return float(user_input) if user_input else default_value

def collect_user_inputs():
    slab_levels = []
    num_slab_levels = int(input("Enter the number of slab levels (default 4): ") or "4")
    for i in range(num_slab_levels):
        name = input(f"Enter the name for Slab Level {i + 1} (e.g., SLAB_LEVEL_{i + 1}): ") or f"SLAB_LEVEL_{i + 1}"
        level = get_input(f"Enter the level (z-coordinate in meters) for {name} (default {3 * (i + 1)}m): ", 3 * (i + 1))
        thickness = get_input(f"Enter the thickness for {name} (in mm, default 200mm): ", 200)
        slab_levels.append({"name": name, "level": level, "thickness": thickness})

    wall_details = {
        "left": {
            "toe_level": get_input("Enter the toe level for the left wall (z-coordinate in meters, default 0m): ", 0),
            "top_level": get_input("Enter the top level for the left wall (z-coordinate in meters, default 15m): ", 15),
            "thickness": get_input("Enter the thickness for the left wall (in mm, default 300mm): ", 300)
        },
        "right": {
            "toe_level": get_input("Enter the toe level for the right wall (z-coordinate in meters, default 0m): ", 0),
            "top_level": get_input("Enter the top level for the right wall (z-coordinate in meters, default 15m): ", 15),
            "thickness": get_input("Enter the thickness for the right wall (in mm, default 300mm): ", 300)
        }
    }

    section_width = get_input("Enter the section width (in meters, default 12m): ", 12)

    return slab_levels, wall_details, section_width

def create_nodes_and_bars(robot, slab_levels, wall_details, section_width):
    try:
        geometry = robot.Project.Structure.Nodes
        bars = robot.Project.Structure.Bars
        node_id = 1

        # Create nodes for each slab level
        for slab in slab_levels:
            geometry.Create(node_id, 0, 0, slab["level"])  # Start node
            geometry.Create(node_id + 1, section_width, 0, slab["level"])  # End node
            bars.Create(node_id, node_id, node_id + 1)  # Create bar between nodes
            node_id += 2  # Increment for next pair of nodes

        # Create nodes and bars for the walls
        left_wall_base_id = node_id
        right_wall_base_id = node_id + 2
        geometry.Create(left_wall_base_id, 0, 0, wall_details["left"]["toe_level"])
        geometry.Create(left_wall_base_id + 1, 0, 0, wall_details["left"]["top_level"])
        bars.Create(left_wall_base_id, left_wall_base_id, left_wall_base_id + 1)

        geometry.Create(right_wall_base_id, section_width, 0, wall_details["right"]["toe_level"])
        geometry.Create(right_wall_base_id + 1, section_width, 0, wall_details["right"]["top_level"])
        bars.Create(right_wall_base_id, right_wall_base_id, right_wall_base_id + 1)

        print("Nodes and bars created successfully!")
    except Exception as e:
        print(f"Failed to create nodes and bars: {e}")

def main():
    Robot = initialize_robot_frame_2d()
    slab_levels, wall_details, section_width = collect_user_inputs()
    create_nodes_and_bars(Robot, slab_levels, wall_details, section_width)

if __name__ == "__main__":
    main()
