#!/bin/bash

# Fix all result.error || occurrences with 'error' in result ? result.error :

# List of files to fix
files=(
  "app/hod/class/[classId]/_components/pupil-form.tsx"
  "app/hod/subject/[subjectId]/_components/edit-class-form.tsx"
  "app/hod/subject/[subjectId]/_components/class-form.tsx"
  "app/hod/subject/[subjectId]/_components/group-form.tsx"
  "app/hod/subject/[subjectId]/group/[groupId]/_components/comment-form.tsx"
  "app/hod/subject/[subjectId]/_components/edit-group-form.tsx"
  "app/hod/subject/[subjectId]/_components/teacher-assignment.tsx"
  "app/admin/_components/create-user-form.tsx"
  "app/admin/_components/pupil-management.tsx"
)

for file in "${files[@]}"; do
  if [ -f "$file" ]; then
    echo "Fixing $file..."
    sed -i '' "s/result\.error || /\'error\' in result ? result.error : /g" "$file"
  fi
done

echo "Done!"
