package main

import (
	"fmt"
	"os"
)

func main() {
	fmt.Fprintln(os.Stdout, os.Args)
	if len(os.Args) > 1 {
		if os.Args[1] == "v2" {
			v2(os.Args[1:])
		} else if os.Args[1] == "v1" {
			v1(os.Args[1:])
		} else {
			v1(os.Args)
		}
	} else {
		v1(os.Args)
	}
}
