#!/bin/bash
block_size_default=100

if [ -n "$1" ]; then
  block_size="$1"
else
  block_size="${block_size_default}"
fi

while true; do
  read -p "Enter block-size [${block_size}]: " input

  if [ -z "${input}" ]; then
    # user pressed <return> with no input
    break
  elif [[ "${input}" =~ ^[0-9]+$ ]]; then
    # user entered a valid number
    block_size="${input}"
    read -p "Okay to proceed with block-size set to ${block_size}? [y/n] " confirmation
    if [ "${confirmation}" = "y" ]; then
      break
    fi
  fi
done

# rest of program goes here, using the chosen block-size

echo "python3 getser.py snos-from-jerry.xlsx ${block_size}"
