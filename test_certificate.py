from certificate_agent.tools import (
    read_student_data,
    generate_certificate
)


students = read_student_data()

print(f"Found {len(students)} students.")

for student in students:
    print(f"Generating certificate for: {student.get('Name', 'Unknown')}")

    certificate = generate_certificate(student)

    print(f"Created: {certificate}")

print()
print("All certificates generated successfully.")