#!/bin/bash
# Package the two application modules and deploy the existing CloudFormation stack.
# Runtime configuration, tokens, and event state are uploaded separately.
set -euo pipefail

if [[ $# -ne 4 ]]; then
    echo "Usage: bash aws/deploy.sh PROFILE BUCKET STACK DISABLED|ENABLED" >&2
    exit 2
fi

aws_profile=$1
sync_bucket=$2
stack_name=$3
schedule_state=$4
case "$schedule_state" in
    DISABLED|ENABLED) ;;
    *) echo "Schedule state must be DISABLED or ENABLED." >&2; exit 2 ;;
esac

repo_root=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.." && pwd)
cd "$repo_root"
build_dir=$(mktemp -d)
trap 'rm -rf -- "$build_dir"' EXIT

# Export the lock without resolving or upgrading packages. All build files are temporary.
uv export --frozen --no-dev --no-emit-project --no-editable \
    --output-file "$build_dir/requirements.txt"
uv pip install --python 3.13 --python-platform x86_64-manylinux2014 \
    --only-binary :all: --require-hashes --no-compile-bytecode \
    --target "$build_dir/package" --requirements "$build_dir/requirements.txt"
cp src/outlook_to_google.py src/aws_handler.py "$build_dir/package/"
uv run --frozen --no-sync python -c \
    'import shutil, sys; shutil.make_archive(sys.argv[1], "zip", sys.argv[2])' \
    "$build_dir/function" "$build_dir/package"

# A content-based key lets CloudFormation detect each distinct code package.
package_hash=$(sha256sum "$build_dir/function.zip")
package_hash=${package_hash%% *}
code_key="deployment/${package_hash}.zip"
aws --profile "$aws_profile" --region us-west-2 s3 cp \
    "$build_dir/function.zip" "s3://${sync_bucket}/${code_key}" --only-show-errors
aws --profile "$aws_profile" --region us-west-2 cloudformation deploy \
    --template-file aws/template.yaml --stack-name "$stack_name" \
    --capabilities CAPABILITY_IAM --no-fail-on-empty-changeset \
    --parameter-overrides "RuntimeBucket=${sync_bucket}" "CodeKey=${code_key}" \
    "ScheduleState=${schedule_state}"
aws --profile "$aws_profile" --region us-west-2 cloudformation describe-stacks \
    --stack-name "$stack_name" --query 'Stacks[0].Outputs' --output table
