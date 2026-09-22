def check_cassandra():
    print("\n========== CASSANDRA ==========\n")

    config = load_key_value_file(CASSANDRA_FILE)
    results = []

    for name, values in config.items():

        server = values.get("server", "").strip()
        command = values.get(
            "command",
            "nodetool status"
        ).strip()

        ssh_result = execute_ssh_command(
            server,
            command
        )

        status = "No"
        details = ""
        nodetool_output = ""

        if ssh_result["success"]:

            # Complete nodetool status output
            nodetool_output = ssh_result["output"]

            # Check whether all Cassandra nodes are UN
            status_bool, details = cassandra_is_all_un(
                nodetool_output
            )

            status = "Yes" if status_bool else "No"

        else:

            details = ssh_result["error"]

        results.append({
            "name": name,
            "server": server,
            "command": command,
            "status": status,
            "details": details,
            "nodetool_output": nodetool_output
        })

        print(
            f"{name}: {status} | {details}"
        )

    return results
