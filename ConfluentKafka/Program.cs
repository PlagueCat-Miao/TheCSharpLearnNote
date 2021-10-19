using System;
using Confluent.Kafka;
using System.Threading;

namespace ConfluentKafka
{
    class Program
    {
        static void Main(string[] args)
        {
            //var p = new Producer();
            //p.Fire();
            Console.WriteLine("GoConsumer");
            GoConsumer();
            
        }
        public static void GoConsumer()
        {
            var conf = new ConsumerConfig
            {
                GroupId = "test-consumer-group",
                BootstrapServers = "localhost:9092",
                // Note: The AutoOffsetReset property determines the start offset in the event
                // there are not yet any committed offsets for the consumer group for the
                // topic/partitions of interest. By default, offsets are committed
                // automatically, so in this example, consumption will only start from the
                // earliest message in the topic 'my-topic' the first time you run the program.
                AutoOffsetReset = AutoOffsetReset.Earliest
            };

            using (var c = new ConsumerBuilder<Ignore, string>(conf).Build())
            {
                c.Subscribe("my-topic");

                CancellationTokenSource cts = new CancellationTokenSource();
                Console.CancelKeyPress += (_, e) => {
                    e.Cancel = true; // prevent the process from terminating.
                    cts.Cancel();
                };

                try
                {
                    while (true)
                    {
                        try
                        {
                            var cr = c.Consume(cts.Token);
                            Console.WriteLine($"Consumed message '{cr.Value}' at: '{cr.TopicPartitionOffset}'.");
                        }
                        catch (ConsumeException e)
                        {
                            Console.WriteLine($"Error occured: {e.Error.Reason}");
                        }
                    }
                }
                catch (OperationCanceledException)
                {
                    // Ensure the consumer leaves the group cleanly and final offsets are committed.
                    c.Close();
                }
            }
        }
    }

    class Producer
    {

        IProducer<Null, string> producer = null;
        public Producer()
        {
            var conf = new ProducerConfig { BootstrapServers = "localhost:9092" };
            producer = new ProducerBuilder<Null, string>(conf).Build();
        }
        void handlerFunc2(DeliveryReport<Null, string> r)
        {

        }

        public void Fire()
        {
            Action<DeliveryReport<Null, string>> handlerFunc1 = r =>
                Console.WriteLine(!r.Error.IsError
                    ? $"Delivered message to {r.TopicPartitionOffset}"
                    : $"Delivery Error: {r.Error.Reason}");

            for (int i = 0; i < 44; ++i)
            {
                producer.Produce("my-topic", new Message<Null, string> { Value = i.ToString() }, handlerFunc1);
            }

            // wait for up to 10 seconds for any inflight messages to be delivered.
            producer.Flush(TimeSpan.FromSeconds(10));

        }
    }

}



