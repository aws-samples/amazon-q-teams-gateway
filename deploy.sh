#!/bin/bash

# Exit on error
set -e

echo "Running cdk deploy..."
cdk deploy -c environment=./environment.json --outputs-file ./cdk-outputs.json


echo "All done!"