CC = g++

CFLAGS = -I ../libxl-4.0.3.1/include_cpp -L ../libxl-4.0.3.1/lib64 -lxl -Wl,-rpath,../libxl-4.0.3.1/lib64

main: main.cpp
	$(CC) -o main main.cpp $(CFLAGS)

clean:
	rm -f write read edit custom format invoice performance *.xls *.xlsx

