/**   Trainer 6502 v0.1
 *    Copyright (C) 2014 Joey Shepard
 *
 *    This program is free software: you can redistribute it and/or modify
 *    it under the terms of the GNU General Public License as published by
 *    the Free Software Foundation, either version 3 of the License, or
 *    (at your option) any later version.
 *
 *    This program is distributed in the hope that it will be useful,
 *    but WITHOUT ANY WARRANTY; without even the implied warranty of
 *    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *    GNU General Public License for more details.
 *
 *    You should have received a copy of the GNU General Public License
 *    along with this program.  If not, see <http://www.gnu.org/licenses/>.
**/

#include <msp430.h>
#include <stdbool.h>
#include <string.h>

#define LED                 BIT0  //P1.0
#define UART_RXD            BIT1  //P1.1
#define UART_TXD            BIT2  //P1.2
#define IO_DATA_CS          BIT3  //P1.3 IO chip 2 select
#define IO_ADDRESS_CS       BIT4  //P1.4 IO chip 1 select
#define IO_CLOCK            BIT5  //P1.5 IO clock
#define IO_MISO             BIT6  //P1.6 IO data out
#define IO_MOSI             BIT7  //P1.7 IO data in

//Outputs (MCU to CPU)
#define CPU_RDY             BIT0  //P2.0 CPU RDY pin 2
#define CPU_ABORT           BIT1  //P2.1 CPU ABORT pin 3
#define CPU_IRQ             BIT2  //P2.2 CPU IRQ pin 4
#define RAM_CS              BIT3  //P2.3 SRAM chip select
#define CPU_BE              BIT4  //P2.4 CPU BE pin 36
#define CPU_PHI2            BIT5  //P2.5 CPU PHI2 pin 37
#define CPU_NMIB            BIT6  //P2.6 CPU NMIB pin 6
#define CPU_RESB            BIT7  //P2.7 CPU RESB pin 40

//Inputs (CPU to IO to MCU)
#define UNUSED2             BIT0 //Button
#define CPU_VPB             BIT1
#define CPU_MLB             BIT2
#define CPU_VPA             BIT3
#define CPU_VDA             BIT4
#define CPU_MX              BIT5
#define CPU_E               BIT6
#define CPU_RWB             BIT7

#define OP                  0x40
#define READ                1
#define WRITE               0

#define PING                10
#define PONG                15
#define ERROR               20
#define UNKNOWN             25
#define RESET_CPU           30
#define RESET_ACK           35
#define DOWN_CYCLE          40
#define UP_CYCLE            45
#define DOWN_ACK            50
#define UP_ACK_READ         55
#define UP_ACK_WRITE        60
#define UPDATE_RAM          65
#define UPDATE_RAM_ACK      70
#define UPDATE_RAM_CRC      75
#define GET_RAM_CRC         80
#define SEND_RAM_CRC        85
#define BEGIN_EMULATING     90
#define UPDATE_DIRTY_RAM    95
#define STOP_EMULATING      100
#define KEEP_EMULATING      105
#define COMM_CHECK          110
#define COMM_CHECK_ACK      115
#define CUSTOM_CHECK        120
#define CUSTOM_CHECK_ACK    125

#define AttribReadonly      0x01
#define AttribCode          0x02
#define AttribBreakpoint    0x04
#define AttribUninitialized 0x80

#define IODIRA              0x00
#define IODIRB              0x01
#define IOCON               0x0A
#define GPPUA               0x0C
#define GPPUB               0x0D
#define GPIOA               0x12
#define GPIOB               0x13
#define OLATA               0x14
#define OLATB               0x15

#define RAM_READ            0x03
#define RAM_WRITE           0x02

#define PACKET_SIZE         20
#define PACKET_END          PACKET_SIZE-1

#define COMPACT_MODE

void UART_Send(unsigned char data);
void UART_Text(char *data);
unsigned char UART_Receive(int timeout);
void UART_Hex(unsigned char data);

unsigned char SPI_Send(unsigned char data);
void SPI_Text(unsigned char *data);

void IO_Data_Send(unsigned char address, unsigned char data);
void IO_Address_Send(unsigned char address, unsigned char data);
unsigned char IO_Data_Get(unsigned char address);
unsigned char IO_Address_Get(unsigned char address);

void delay_ms(int ms);

bool Wait_Command();
void Send_Command();

void debug(int ontime, int offtime);

unsigned char in_buffer[PACKET_SIZE];
unsigned char out_buffer[PACKET_SIZE];
volatile bool UART_Failed,reset;

//changed this from known working version!
//void main(void)
int main(void)
{
  WDTCTL=WDTPW + WDTHOLD;

  BCSCTL1=CALBC1_16MHZ;
  DCOCTL=CALDCO_16MHZ;

  BCSCTL3|=LFXT1S_2;

  TA0CCR0=12000;
  TA0CCTL0=CCIE;
  TA0CTL=MC_0|ID_0|TASSEL_1|TACLR;

  TA1CCR0=2400;//every 200ms
  //TA1CCTL0=CCIE;
  TA1CTL=MC_0|ID_0|TASSEL_1|TACLR;

  UCA0CTL1=UCSWRST|UCSSEL_2;
  UCA0CTL0 = 0;
  //9.6k
  //UCA0MCTL = UCBRS_5+UCBRF_0;
  //UCA0BR0 = 0x82;
  //UCA0BR1 = 0x06;

  //57.6k
  //UCA0MCTL = UCBRS_6+UCBRF_0;
  //UCA0BR0 = 0x15;
  //UCA0BR1 = 0x01;

  //115.2k
  //UCA0MCTL = UCBRS_7+UCBRF_0;
  //UCA0BR0 = 0x8A;
  //UCA0BR1 = 0x00;

  //500k
  //UCA0MCTL = UCBRS_0+UCBRF_0;
  //UCA0BR0 = 0x20;
  //UCA0BR1 = 0x00;

  //Comment this out?
  //1000k
  UCA0MCTL = UCBRS_0+UCBRF_0;
  UCA0BR0 = 0x10;
  UCA0BR1 = 0x00;

  #ifdef COMPACT_MODE
  //500k
  UCA0MCTL = UCBRS_0+UCBRF_0;
  UCA0BR0 = 0x20;
  UCA0BR1 = 0x00;
  #else
  //1000k
  UCA0MCTL = UCBRS_0+UCBRF_0;
  UCA0BR0 = 0x10;
  UCA0BR1 = 0x00;
  #endif

  UCA0CTL1&=~UCSWRST;

  UCB0CTL1=UCSWRST;
  UCB0CTL0=UCCKPH|UCMST|UCSYNC|UCMSB;//mode 0
  UCB0CTL1|=UCSSEL_2;
  //UCB0BR0=100;//160khz?
  UCB0BR0=2;//would 1 work?
  UCB0BR1=0;
  UCB0CTL1&=~UCSWRST;

  P1OUT=IO_DATA_CS|IO_ADDRESS_CS;
  P1DIR=LED|IO_DATA_CS|IO_ADDRESS_CS;
  P1SEL= IO_CLOCK|IO_MISO|IO_MOSI|UART_RXD|UART_TXD;
  P1SEL2=IO_CLOCK|IO_MISO|IO_MOSI|UART_RXD|UART_TXD;

  P2OUT=CPU_RDY|CPU_ABORT|CPU_IRQ|CPU_BE|CPU_PHI2|CPU_NMIB|CPU_RESB|RAM_CS;
  P2DIR=CPU_RDY|CPU_ABORT|CPU_IRQ|CPU_BE|CPU_PHI2|CPU_NMIB|CPU_RESB|RAM_CS;
  P2SEL=0;
  P2SEL2=0;

  /*P2OUT=RESET_BUTTON;
  P2REN=RESET_BUTTON;
  P2IE|=RESET_BUTTON;
  P2IES|=RESET_BUTTON;//falling edge?
  P2IFG&=~RESET_BUTTON;*/

  unsigned char IOctr;
  IO_Data_Send(IOCON, 0x18);
  IO_Data_Send(IODIRA, 0xFF);
  for (IOctr=1;IOctr<0xB;IOctr++) IO_Data_Send(IOctr,0);
  IO_Data_Send(IODIRB, 0xFF);
  for (IOctr=11;IOctr<0x1B;IOctr++) IO_Data_Send(IOctr,0);
  //IO_Data_Send(GPPUA, 0x00);
  //IO_Data_Send(GPPUB, 0x00);


  IO_Address_Send(IOCON, 0x18);
  IO_Address_Send(IODIRA, 0xFF);
  for (IOctr=1;IOctr<0xB;IOctr++) IO_Address_Send(IOctr,0);
  IO_Address_Send(IODIRB, 0xFF);
  for (IOctr=11;IOctr<0x1B;IOctr++) IO_Address_Send(IOctr,0);
  //IO_Address_Send(GPPUA, 0x00);
  //IO_Address_Send(GPPUB, 0x00);

  __enable_interrupt();

  //these probably don't have to be volatile
  volatile unsigned int i,j,k;
  volatile unsigned int ptr;
  volatile unsigned char crc,buff;

  unsigned long new_address;
  unsigned char new_data, new_flags, new_status;
  bool emulating=false;
  unsigned int DirtyAddress[64];
  unsigned char DirtyData[64];
  unsigned int DirtyCount;
  bool DirtyFound;
  unsigned int CycleCount;
  unsigned int ExitMode;
  unsigned int ExitFlags;
  unsigned int counter;
  bool ResetInputs;

  for (buff=0;buff<2;buff++)
  {
    for (i=0;i<600;i+=((i/5)+1))
    {
      for (j=0;j<60;j++)
      {
        P1OUT|=LED;
        for (k=0;k<i;k++);
        P1OUT&=~LED;
        for (k=i;k<1000;k++);
      }
    }
    for (i=600;i<=600;i-=((i/5)+1))
    {
      for (j=0;j<60;j++)
      {
        P1OUT|=LED;
        for (k=0;k<i;k++);
        P1OUT&=~LED;
        for (k=i;k<1000;k++);
      }
    }
  }

  while(1)
  {
    //reset code
    reset=false;
    while (!reset)
    {
      P1OUT&=~LED;
      if (!Wait_Command())
      {
        debug(300,0);
      }
      else
      {
        P1OUT|=LED;
        switch (in_buffer[PACKET_END])
        {
          case PING: //ping
            if (memcmp(in_buffer,"PING!",5))
            {
              out_buffer[PACKET_END]=ERROR;
              out_buffer[PACKET_END-1]=PING;
              memcpy(out_buffer,"ERROR   ",8);
              Send_Command();
            }
            else
            {
              out_buffer[PACKET_END]=PONG;
              memcpy(out_buffer,"PONG!",5);
              Send_Command();
            }
            break;
          case RESET_CPU:
            IO_Data_Send(IODIRA, 0xFF);

            P2OUT|=CPU_RESB;
            P2OUT|=CPU_PHI2;
            P2OUT&=~CPU_RESB;

            for (i=0;i<10;i++) {P2OUT&=~CPU_PHI2;P2OUT|=CPU_PHI2;}

            P2OUT|=CPU_RESB;

            for (i=0;i<6;i++) {P2OUT&=~CPU_PHI2;P2OUT|=CPU_PHI2;}

            out_buffer[PACKET_END]=RESET_ACK;
            Send_Command();

            break;
          case DOWN_CYCLE:
            P2OUT&=~CPU_PHI2;
            IO_Data_Send(IODIRA, 0xFF);//set to inputs
            out_buffer[PACKET_END]=DOWN_ACK;
            out_buffer[0]=IO_Address_Get(GPIOA);
            out_buffer[1]=IO_Address_Get(GPIOB);
            out_buffer[2]=IO_Data_Get(GPIOB);//flags
            Send_Command();
            break;
          case UP_CYCLE:
            P2OUT|=CPU_PHI2;

            //out_buffer was set in DOWN_CYCLE
            if (out_buffer[2]&CPU_RWB)//read
            {
              IO_Data_Send(OLATA,in_buffer[0]);
              IO_Data_Send(IODIRA, 0x00);//set to outputs
              out_buffer[3]=0;
              out_buffer[PACKET_END]=UP_ACK_READ;

              if (in_buffer[12]==0xAA)//free running
              {
                //Save address
                ptr=(out_buffer[0]<<8)+out_buffer[1];

                for (i=1;i<12;i++)
                {
                  //Down cycle
                  P2OUT&=~CPU_PHI2;
                  out_buffer[0]=IO_Address_Get(GPIOA);
                  out_buffer[1]=IO_Address_Get(GPIOB);
                  out_buffer[2]=IO_Data_Get(GPIOB);//status
                  //Check to see if new address is in range
                  if ((out_buffer[2]&CPU_RWB)&&((out_buffer[0]<<8)+out_buffer[1]==ptr+i))
                  {
                    out_buffer[3]++;

                    //Up cycle
                    P2OUT|=CPU_PHI2;
                    IO_Data_Send(OLATA,in_buffer[i]);
                  }
                  else i=100;
                }
              }//free running
            }
            else//write
            {
              out_buffer[3]=IO_Data_Get(GPIOA);
              out_buffer[PACKET_END]=UP_ACK_WRITE;
            }
            Send_Command();
            break;
          case UPDATE_RAM:
            P2OUT&=~RAM_CS;
            SPI_Send(RAM_WRITE);
            SPI_Send(0);
            SPI_Send(0);
            SPI_Send(0);
            out_buffer[PACKET_END]=UPDATE_RAM_ACK;
            Send_Command();

            //counter=0;
            for (j=0;j<5;j++)
            {
              crc=0;
              //counter=0;

              #ifdef COMPACT_MODE
              if (j==4) k=0x2000;
              else k=0x3800;
              for (i=0;i<k;i++)
              {
                buff=UART_Receive(500);
                //counter++;
                if (UART_Failed) for (;;) debug(500,100);
                SPI_Send(buff);
                crc+=buff;
                if (buff&AttribUninitialized) buff=0;
                else
                {
                  counter++;
                  buff=UART_Receive(500);
                }
                //Do something more meaningful
                if (UART_Failed) for (;;) debug(100,500);
                SPI_Send(buff);
                crc+=buff;
              }
              #else
              if (j==4) k=0x4000;
              else k=0x7000;

              for (i=0;i<k;i++)
              {
                buff=UART_Receive(500);
                SPI_Send(buff);
                //counter++;
                crc+=buff;
                if (UART_Failed)
                {
                  //Hard failure while updating RAM
                  for (;;) debug(500,100);
                }
              }
              #endif

              //Not checking CRC???
              //UART_Receive(0);
              //if (UART_Receive(0)!=crc) for (;;) debug(100,100);

              out_buffer[0]=crc;
              out_buffer[1]=UART_Receive(0);
              //out_buffer[2]=(counter&0xFF);
              //out_buffer[3]=counter>>8;
              out_buffer[PACKET_END]=UPDATE_RAM_CRC;
              Send_Command();
            }

            P2OUT|=RAM_CS;
            break;
          case GET_RAM_CRC:
            P2OUT&=~RAM_CS;
            SPI_Send(RAM_READ);
            SPI_Send(0);
            SPI_Send(0);
            SPI_Send(0);
            for (j=0;j<5;j++)
            {
              crc=0;
              if (j==4) k=0x4000;
              else k=0x7000;
              for (i=0;i<k;i++)
              {
                crc+=SPI_Send(0);
              }
              out_buffer[j]=crc;
            }
            out_buffer[PACKET_END]=SEND_RAM_CRC;
            Send_Command();
            P2OUT|=RAM_CS;
            break;
          case BEGIN_EMULATING:
            DirtyCount=0;
            CycleCount=0;
            ExitMode=0;
            ExitFlags=in_buffer[0];
            emulating=true;
            ResetInputs=true;
            TA1CCR0=in_buffer[1]+(in_buffer[2]<<8);
            TA1CTL=MC_1|ID_0|TASSEL_1|TACLR;

            while(emulating)
            {
            //case DOWN_CYCLE:
              P2OUT&=~CPU_PHI2;
              //maybe dont do this unless really necessary
              if (ResetInputs)
              {
                IO_Data_Send(IODIRA, 0xFF);//set to inputs
                ResetInputs=false;
              }
              new_status=IO_Data_Get(GPIOB);//status

              if (new_status & (CPU_VDA | CPU_VPA))
              {
                new_address=((unsigned long)IO_Address_Get(GPIOB))<<1;//low byte
                new_address+=((unsigned long)IO_Address_Get(GPIOA))<<9;//high byte
              }

              //Send_Command();
            //case UP_CYCLE:
              P2OUT|=CPU_PHI2;
              if (new_status & (CPU_VDA | CPU_VPA))
              {
                if (new_status&CPU_RWB)//read
                {
                  P2OUT&=~RAM_CS;
                  SPI_Send(RAM_READ);
                  SPI_Send(new_address>>16)&0xFF;
                  SPI_Send((new_address>>8)&0xFF);
                  SPI_Send(new_address&0xFF);
                  new_flags=SPI_Send(0);
                  new_data=SPI_Send(0);
                  P2OUT|=RAM_CS;
                  IO_Data_Send(OLATA,new_data);
                  IO_Data_Send(IODIRA, 0x00);//set to outputs
                  ResetInputs=true;
                  if ((new_flags&AttribUninitialized)&ExitFlags) ExitMode|=AttribUninitialized;
                  if (new_flags&AttribBreakpoint) ExitMode|=AttribBreakpoint;
                }
                else//write
                {
                  new_data=IO_Data_Get(GPIOA);
                  P2OUT&=~RAM_CS;
                  SPI_Send(RAM_READ);
                  SPI_Send(new_address>>16);
                  SPI_Send((new_address>>8)&0xFF);
                  SPI_Send(new_address&0xFF);
                  new_flags=SPI_Send(0);
                  P2OUT|=RAM_CS;

                  if ((new_flags&AttribReadonly)&ExitFlags)
                  {
                    ExitMode|=AttribReadonly;
                  }
                  else
                  {
                    P2OUT&=~RAM_CS;
                    SPI_Send(RAM_WRITE);
                    SPI_Send(new_address>>16);
                    SPI_Send((new_address>>8)&0xFF);
                    SPI_Send(new_address&0xFF);
                    SPI_Send(new_flags&(~AttribUninitialized));
                    SPI_Send(new_data);
                    P2OUT|=RAM_CS;
                    //check here for multiplier
                  }


                  DirtyFound=false;
                  for (i=0;i<DirtyCount;i++)
                  {
                    if (DirtyAddress[i]==new_address>>1)
                    {
                      DirtyData[i]=new_data;
                      DirtyFound=true;
                      i=1000;
                    }
                  }
                  if (DirtyFound==false)
                  {
                    DirtyAddress[DirtyCount]=new_address>>1;
                    DirtyData[DirtyCount]=new_data;
                    DirtyCount++;
                  }
                }
              }
              else
              {
                //what to do if bus unused
                //seems to add about 10% performance
              }
              CycleCount++;

              if ((TA1CTL&TAIFG)||(DirtyCount==64)||(ExitMode))
              {
                if (TA1CTL&TAIFG) TA1CTL&=~TAIFG;
                else TA1CTL=MC_1|ID_0|TASSEL_1|TACLR;

                UART_Send(UPDATE_DIRTY_RAM);
                UART_Send(ExitMode);//Extra flags like breakpoint
                UART_Send((new_address>>1)&0xFF);//Flag address
                UART_Send(new_address>>9);//Flag address
                UART_Send(CycleCount&0xFF);//Cycle count
                UART_Send(CycleCount>>8);//Cycle count
                UART_Send(DirtyCount);

                for (i=0;i<DirtyCount;i++)
                {
                  UART_Send(DirtyAddress[i]&0xFF);//Could take out whiles to speed up
                  UART_Send(DirtyAddress[i]>>8);
                  UART_Send(DirtyData[i]);
                }
                DirtyCount=0;
                CycleCount=0;
                j=UART_Receive(500);
                if (UART_Failed) emulating=false;
                if (emulating)
                {
                  if (j==KEEP_EMULATING)
                  {
                    k=UART_Receive(500);
                    for (i=0;i<k;i++)
                    {
                      new_address=((unsigned long)UART_Receive(500))<<1;//low byte
                      new_address+=((unsigned long)UART_Receive(500))<<9;//high byte
                      new_address++;
                      P2OUT&=~RAM_CS;
                      SPI_Send(RAM_WRITE);
                      SPI_Send(new_address>>16);
                      SPI_Send((new_address>>8)&0xFF);
                      SPI_Send(new_address&0xFF);
                      SPI_Send(UART_Receive(500));
                      P2OUT|=RAM_CS;
                    }
                  }
                  else if (j==STOP_EMULATING)
                  {
                    UART_Send(STOP_EMULATING);
                    TA1CTL=MC_0|ID_0|TASSEL_1|TACLR;
                    emulating=false;
                  }
                }
              }
            }

            break;
          case COMM_CHECK:
            out_buffer[PACKET_END]=COMM_CHECK_ACK;
            memcpy(out_buffer,in_buffer,16);
            Send_Command();
            break;
          case CUSTOM_CHECK:
            IO_Data_Send(IODIRA, 0xFF);//set to inputs
            out_buffer[0]=1;
            out_buffer[1]=IO_Address_Get(GPIOA);
            out_buffer[2]=IO_Address_Get(GPIOB);
            out_buffer[5]=0;
            for (i=0;i<5000;i++)
            {
              out_buffer[3]=IO_Address_Get(GPIOA);
              out_buffer[4]=IO_Address_Get(GPIOB);
              out_buffer[5]++;
              if ((out_buffer[1]!=out_buffer[3])||(out_buffer[2]!=out_buffer[4]))
              {
                out_buffer[0]=0;
                i=30000;
              }
              delay_ms(1);
            }
            Send_Command();
            break;
          default:
            out_buffer[PACKET_END]=UNKNOWN;
            out_buffer[PACKET_END-1]=in_buffer[PACKET_END];
            memcpy(out_buffer,"UNKNOWN  ",9);
            Send_Command();
            break;
        }
      }
    }
  }
}

__attribute__((interrupt(TIMER0_A0_VECTOR))) static void TA0_ISR(void)
{
  P1OUT|=LED;
  delay_ms(10);
  P1OUT&=~LED;
  delay_ms(10);
  UART_Failed=true;
  TA0CTL=MC_0|ID_0|TASSEL_1|TACLR;
}

__attribute__((interrupt(TIMER1_A0_VECTOR))) static void TA1_ISR(void)
{
  //if (test) test=0;
  //else test=1;
}
__attribute__((interrupt(PORT2_VECTOR))) static void P2_ISR(void)
{
  reset=true;
  //P2IFG&=~RESET_BUTTON;
}

void UART_Send(unsigned char data)
{
  while(!(UC0IFG&UCA0TXIFG));
  UCA0TXBUF=data;
  while (UCA0STAT & UCBUSY);
}

void UART_Text(char *data)
{
  int i=0;
  while (data[i]) UART_Send(data[i++]);
}

unsigned char UART_Receive(int timeout)
{
  UART_Failed=false;

  if (timeout)
  {
    //what is this?
    if (TA0CTL&TAIFG) debug(100,500);
    //TA0CTL&=~TAIFG;
    TA0CCR0=timeout*12;
    TA0CTL=MC_1|ID_0|TASSEL_1|TACLR;
  }

  if (UCA0STAT&UCOE)
  {
    UART_Failed=true;
    debug(1000,2000);
    return 0;
  }

  while ((!(UC0IFG&UCA0RXIFG)))
  {
    if (UART_Failed)//time out
    {
      debug(1000,1000);
      debug(1000,1000);
      return 0;
    }
  }
  if (timeout) TA0CTL=MC_0|ID_0|TASSEL_1|TACLR;
  return UCA0RXBUF;
}

void UART_Hex(unsigned char data)
{
  unsigned char buff;
  buff=data/16;
  if (buff>9) buff+=55;
  else buff+='0';
  UART_Send(buff);
  buff=data%16;
  if (buff>9) buff+=55;
  else buff+='0';
  UART_Send(buff);
}

unsigned char SPI_Send(unsigned char data)
{
  unsigned char buff;
  while(!(UC0IFG&UCB0TXIFG));
  UCB0TXBUF=data;
  while (UCB0STAT & UCBUSY);

  buff=UCB0RXBUF;
  return buff;
}

void SPI_Text(unsigned char *data)
{
  int i=0;
  while (data[i])
  {
    SPI_Send(data[i]);
    i++;
  }
  return;
}

void IO_Data_Send(unsigned char address, unsigned char data)
{
  P1OUT&=~IO_DATA_CS;
  SPI_Send(OP | WRITE);
  SPI_Send(address);
  SPI_Send(data);
  P1OUT|=IO_DATA_CS;
}

void IO_Address_Send(unsigned char address, unsigned char data)
{
  P1OUT&=~IO_ADDRESS_CS;
  SPI_Send(OP | WRITE);
  SPI_Send(address);
  SPI_Send(data);
  P1OUT|=IO_ADDRESS_CS;
}

unsigned char IO_Data_Get(unsigned char address)
{
  unsigned char buff;
  P1OUT&=~IO_DATA_CS;
  SPI_Send(OP | READ);
  SPI_Send(address);
  buff=SPI_Send(0);
  P1OUT|=IO_DATA_CS;
  return buff;
}

unsigned char IO_Address_Get(unsigned char address)
{
  unsigned char buff;
  P1OUT&=~IO_ADDRESS_CS;
  SPI_Send(OP | READ);
  SPI_Send(address);
  buff=SPI_Send(0);
  P1OUT|=IO_ADDRESS_CS;
  return buff;
}

void delay_ms(int ms)
{
  while (ms--) __delay_cycles(16000);
}

bool Wait_Command()
{
  int i,j;
  i=UCA0RXBUF;
  in_buffer[0]=UART_Receive(0);
  for (i=1;i<PACKET_SIZE;i++)
  {
    in_buffer[i]=UART_Receive(500);
    if (UART_Failed)
    {
      /*while(1)
      {
        for (j=0;j<(i+1);j++)
        {
          debug(100,500);
        }
        debug(0,1000);
      }*/
      return false;
    }
  }
  return true;
}

void Send_Command()
{
  int i;
  for (i=0;i<PACKET_SIZE;i++) UART_Send(out_buffer[i]);
}

void debug(int ontime, int offtime)
{
  P1OUT|=LED;
  delay_ms(ontime);
  P1OUT&=~LED;
  delay_ms(offtime);
}
